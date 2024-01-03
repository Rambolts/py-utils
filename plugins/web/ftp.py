import os
from typing import List, BinaryIO

import ftplib
import paramiko
import paramiko.sftp


class FTPPlugin:
    def __init__(self, hostname: str, username: str = None, password: str = None) -> None:
        self._ftp_server = ftplib.FTP(hostname)
        self._ftp_server.login(username, password)

    def upload_file(self, file_path: str) -> None:
        file_name = os.path.basename(file_path)
        with open(file_path, 'rb') as file:
            self._ftp_server.storbinary(f'STOR {file_name}', file)

    def download_file(self, file_name: str, destination_folder: str) -> None:
        file_path = os.path.join(destination_folder, file_name)
        with open(file_path, 'wb') as file:
            self._ftp_server.retrbinary(f'RETR {file_name}', file.write)

    def rename_file(self, from_name: str, to_name: str) -> None:
        self._ftp_server.rename(from_name, to_name)

    def delete_file(self, file_name: str) -> None:
        self._ftp_server.delete(file_name)

    def get_current_directory(self) -> str:
        return self._ftp_server.pwd()

    def set_current_directory(self, dir_path: str) -> None:
        self._ftp_server.cwd(dir_path)

    def create_directory(self, dir_path: str) -> str:
        dir_path = self._ftp_server.mkd(dir_path)
        return dir_path

    def remove_directory(self, dir_path: str) -> None:
        self._ftp_server.rmd(dir_path)

    def list_files(self) -> None:
        self._ftp_server.dir()

    def get_files_list(self) -> List[str]:
        return self._ftp_server.nlst()

    def disconnect(self) -> None:
        self._ftp_server.quit()


class SFTPPlugin:
    def __init__(self, hostname: str, port: int = 22, username: str = None, password: str = None) -> None:
        self._transport = paramiko.Transport((hostname, port))
        self._transport.set_keepalive(30)
        self._transport.connect(username=username, password=password)
        self._sftp_server = paramiko.SFTPClient.from_transport(self._transport)

    def upload_file(self, file_name: str, destination_path: str = None) -> None:
        if not destination_path:
            destination_path = self.get_current_directory()

        filename = os.path.basename(file_name)
        file_path = os.path.join(destination_path, filename)

        self._sftp_server.put(file_name, file_path)

    def download_file(self, file_name: str, destination_folder: str) -> None:
        filename = os.path.basename(file_name)
        file_path = os.path.join(destination_folder, filename)

        self._sftp_server.get(file_name, file_path)

    def download_large_file(self, file_name: str, destination_folder: str, callback=None) -> None:
        filename = os.path.basename(file_name)
        file_path = os.path.join(destination_folder, filename)

        remote_file_size = self._sftp_server.stat(file_name).st_size

        with self._sftp_server.open(file_name, 'rb') as f_in, open(file_path, 'wb') as f_out:
            SFTPFileDownloader(
                f_in=f_in,
                f_out=f_out,
                callback=callback
            ).download()

        local_file_size = os.path.getsize(file_path)
        if remote_file_size != local_file_size:
            raise IOError(f'file size mismatch: {remote_file_size} != {local_file_size}')

    def rename_file(self, from_name: str, to_name: str) -> None:
        self._sftp_server.rename(from_name, to_name)

    def delete_file(self, file_name: str) -> None:
        self._sftp_server.remove(file_name)

    def get_current_directory(self) -> str:
        cwd = self._sftp_server.getcwd()
        if not cwd:
            return '/'
        return cwd

    def set_current_directory(self, dir_path: str) -> None:
        self._sftp_server.chdir(dir_path)

    def create_directory(self, dir_path: str) -> None:
        self._sftp_server.mkdir(dir_path)

    def remove_directory(self, dir_path: str) -> None:
        self._sftp_server.rmdir(dir_path)

    def get_files_list(self) -> List[str]:
        return self._sftp_server.listdir()
    
    def get_file_attributes(self, filename: str) -> paramiko.SFTPAttributes:
        return self._sftp_server.stat(filename)

    def disconnect(self):
        self._sftp_server.close()


class SFTPFileDownloader:

    DOWNLOAD_MAX_REQUESTS = 48
    DOWNLOAD_MAX_CHUNK_SIZE = 0x8000

    def __init__(self, f_in: paramiko.SFTPFile, f_out: BinaryIO, callback=None):
        self.f_in = f_in
        self.f_out = f_out
        self.callback = callback

        self.requested_chunks = {}
        self.received_chunks = {}
        self.saved_exception = None

    def download(self):
        file_size = self.f_in.stat().st_size
        requested_size = 0
        received_size = 0

        while True:
            while len(self.requested_chunks) + len(self.received_chunks) < self.DOWNLOAD_MAX_REQUESTS and \
                    requested_size < file_size:
                chunk_size = min(self.DOWNLOAD_MAX_CHUNK_SIZE, file_size - requested_size)
                request_id = self._sftp_async_read_request(
                    fileobj=self,
                    file_handle=self.f_in.handle,
                    offset=requested_size,
                    size=chunk_size
                )
                self.requested_chunks[request_id] = (requested_size, chunk_size)
                requested_size += chunk_size

            self.f_in.sftp._read_response()
            self._check_exception()

            while True:
                chunk = self.received_chunks.pop(received_size, None)
                if chunk is None:
                    break
                _, chunk_size, chunk_data = chunk
                self.f_out.write(chunk_data)
                if self.callback is not None:
                    self.callback(chunk_data)

                received_size += chunk_size

            if received_size >= file_size:
                break

            if not self.requested_chunks and len(self.received_chunks) >= self.DOWNLOAD_MAX_REQUESTS:
                raise ValueError('SFTP communication error. The queue with requested file chunks is empty and'
                                 'the received chunks queue is full and cannot be consumed.')

        return received_size

    def _sftp_async_read_request(self, fileobj, file_handle, offset, size):
        sftp_client = self.f_in.sftp

        with sftp_client._lock:
            num = sftp_client.request_number

            msg = paramiko.Message()
            msg.add_int(num)
            msg.add_string(file_handle)
            msg.add_int64(offset)
            msg.add_int(size)

            sftp_client._expecting[num] = fileobj
            sftp_client.request_number += 1

        sftp_client._send_packet(paramiko.sftp.CMD_READ, msg)
        return num

    def _async_response(self, t, msg, num):
        if t == paramiko.sftp.CMD_STATUS:
            try:
                self.f_in.sftp._convert_status(msg)
            except Exception as e:
                self.saved_exception = e
            return
        if t != paramiko.sftp.CMD_DATA:
            raise paramiko.SFTPError('Expected data')
        data = msg.get_string()

        chunk_data = self.requested_chunks.pop(num, None)
        if chunk_data is None:
            return

        # save chunk
        offset, size = chunk_data

        if size != len(data):
            raise paramiko.SFTPError(f'Invalid data block size. Expected {size} bytes, but it has {len(data)} size')
        self.received_chunks[offset] = (offset, size, data)

    def _check_exception(self):
        if self.saved_exception is not None:
            x = self.saved_exception
            self.saved_exception = None
            raise x
