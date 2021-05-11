from __future__ import unicode_literals

import struct
import os
import argparse
import re
import datetime

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as tkMsgBox


def main():
    parser = argparse.ArgumentParser('ymca', description='Yahoo Messenger Chat Archive reader')
    parser.add_argument('path', help='Default profile folder', default=None, type=str)
    args = parser.parse_args()

    ymca = Ymca(args.path)
    ymca.mainloop()


class Ymca(tk.Frame):
    friend_node_tag = 'friend'
    archive_node_tag = 'archive'
    archive_date_format = '%Y/%m/%d'

    tag_timestamp = 'timestamp'
    tag_me = 'me'
    tag_friend = 'friend'

    def __init__(self, profile_dir):
        super().__init__()

        self._username = None
        self._archive_name_regex = re.compile(r'(?P<y>\d\d\d\d)(?P<m>\d\d)(?P<d>\d\d)-.*\.dat')
        self._setup_ui()
        if profile_dir:
            self._try_profile_folder(profile_dir)

    def _setup_ui(self):
        self.master.title('YMCA')
        self.master.geometry('600x400')
        self.master.iconbitmap('res/ym.ico')

        self.style = ttk.Style()
        self.style.theme_use('default')

        self._setup_browse_ui()
        self._setup_msg_ui()

        self.pack(fill=tk.BOTH)

    def _setup_browse_ui(self):
        frame = tk.Frame()
        frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        label = tk.Label(frame, text='YM profile folder')
        label.pack(side=tk.LEFT, padx=5)

        self._profile_path = tk.Entry(frame, state='readonly')
        self._profile_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        button = tk.Button(frame, text='Browse')
        button.configure(command=lambda: self._try_profile_folder(filedialog.askdirectory()))
        button.pack(side=tk.RIGHT)

    def _setup_msg_ui(self):
        frame = tk.PanedWindow(orient=tk.HORIZONTAL)
        frame.pack(side=tk.TOP, anchor=tk.N, fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        self._friend_list = ttk.Treeview(frame, show='tree')
        self._friend_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._friend_list.bind('<<TreeviewOpen>>', self._on_friend_dir_open)
        self._friend_list.bind('<<TreeviewSelect>>', self._on_open_archive)
        frame.add(self._friend_list)

        msg_frame = tk.Frame(frame)
        scrollbar = tk.Scrollbar(msg_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._msg_window = tk.Text(msg_frame, state=tk.DISABLED, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self._msg_window.yview)

        self._msg_window.pack(fill=tk.BOTH, expand=True)
        self._msg_window.tag_config(Ymca.tag_timestamp, foreground='grey')
        self._msg_window.tag_config(Ymca.tag_me, foreground='blue')
        self._msg_window.tag_config(Ymca.tag_friend, foreground='red')
        frame.add(msg_frame)

    def _on_friend_dir_open(self, event):
        # pylint: disable=unused-argument
        selected = self._friend_list.selection()
        if not self._friend_list.tag_has(self.friend_node_tag, selected):
            return # We only expand friend nodes
        item = self._friend_list.item(selected)
        friend_dir, already_expanded = item['values']
        if not already_expanded:
            for child in self._friend_list.get_children(selected):
                self._friend_list.delete(child)
            self._expand_friend_archive_list(selected, friend_dir)

    def _expand_friend_archive_list(self, friend_item, friend_dir):
        idx = 0
        for file_name in os.listdir(friend_dir):
            archive_path = os.path.join(friend_dir, file_name)
            if os.path.isdir(archive_path):
                continue
            m = self._archive_name_regex.match(file_name)
            if not m:
                continue
            archive_date = datetime.date(int(m['y']), int(m['m']), int(m['d'])).strftime(
                self.archive_date_format)
            friend_name = self._friend_list.item(friend_item)['text']
            self._friend_list.insert(
                friend_item,
                idx,
                text=archive_date,
                values=(archive_path, friend_name),
                tags=[self.archive_node_tag])
            idx += 1
        self._friend_list.item(friend_item, values=(friend_dir, 'expanded'))

    def _on_open_archive(self, event):
        #pylint: disable=unused-argument
        selected = self._friend_list.selection()
        if not self._friend_list.tag_has(self.archive_node_tag, selected):
            return
        archive_path, friend_name, = self._friend_list.item(selected)['values']
        self._load_archive(archive_path, friend_name)

    def _load_archive(self, archive_path, friend_name):
        archive = YmArchive(self._username, archive_path)

        self._msg_window.config(state=tk.NORMAL)
        self._msg_window.delete('1.0', tk.END)

        msg_count = len(archive.messages)
        for idx, msg in enumerate(archive.messages):
            self._add_chat_msg(friend_name, msg, is_last_msg=(idx == msg_count - 1))

        self._msg_window.config(state=tk.DISABLED)

    def _add_chat_msg(self, friend_name, msg, is_last_msg):
        timestamp = msg.timestamp.strftime('%H:%M:%S ')
        pos = len(timestamp)
        self._msg_window.insert(tk.END, timestamp)
        self._msg_window.tag_add(Ymca.tag_timestamp, 'current linestart', 'current lineend')

        if msg.is_received:
            sender = friend_name
            sender_tag = Ymca.tag_friend
        else:
            sender = self._username
            sender_tag = Ymca.tag_me
        sender += ' '
        self._msg_window.insert(tk.END, sender)
        self._msg_window.tag_add(
            sender_tag,
            'current linestart + %dc' % pos,
            'current lineend')
        pos += len(sender)

        self._msg_window.insert(tk.END, msg.content + ('' if is_last_msg else '\n'))

    def _try_profile_folder(self, path):
        msg_dir = self._find_message_dir(path, max_depth=3)
        profile_dir = self._profile_dir_from_msg_dir(msg_dir)
        if not profile_dir:
            tkMsgBox.showerror(
                title='Oops',
                message='The given directory is not a valid YM profile directory\n(%s)' % path)
            return

        self._profile_path.configure(state=tk.NORMAL)
        self._profile_path.insert(0, path)
        self._profile_path.configure(state='readonly')
        self._load_msg_dir(msg_dir)
        self._username = os.path.basename(profile_dir)

    def _profile_dir_from_msg_dir(self, msg_dir):
        if not msg_dir:
            return None
        # Profile dir is the grand-parent directory of the message dir
        # (some_path/profile/archive/messages)
        path = os.path.abspath(msg_dir)
        for i in range(2):
            path = os.path.abspath(os.path.join(path, os.pardir))
            if not path or not os.path.isdir(path):
                return None
        return path

    def _find_message_dir(self, path, max_depth):
        if not os.path.isdir(path) or max_depth == 0:
            return None

        # Check if the given folder has at least one valid archive file
        for friend_name in os.listdir(path):
            friend_dir = os.path.join(path, friend_name)
            if not os.path.isdir(friend_dir):
                continue
            for file_name in os.listdir(friend_dir):
                archive_path = os.path.join(friend_dir, file_name)
                if os.path.isdir(archive_path):
                    continue
                if self._archive_name_regex.match(file_name):
                    return path

        for child in os.listdir(path):
            return self._find_message_dir(os.path.join(path, child), max_depth - 1)


    def _load_msg_dir(self, path):
        self._clear_friend_list()

        idx = 0
        for friend_name in os.listdir(path):
            friend_dir = os.path.join(path, friend_name)
            if not os.path.isdir(friend_dir):
                continue

            has_archive = False
            for file_name in os.listdir(friend_dir):
                archive_path = os.path.join(friend_dir, file_name)
                if os.path.isdir(archive_path):
                    continue
                if self._archive_name_regex.match(file_name):
                    has_archive = True
                    break
            self._add_friend(idx, friend_name, friend_dir, has_archive)
            idx += 1

    def _clear_friend_list(self):
        for idx in self._friend_list.get_children():
            self._friend_list.delete(idx)

    def _add_friend(self, idx, friend_name, friend_dir, has_archive):
        self._friend_list.get_children()
        parent = self._friend_list.insert(
            '',
            idx,
            text=friend_name,
            values=(friend_dir, ''),
            tags=[self.friend_node_tag])

        if has_archive:
            # Just create a dummy child entry to reduce the load of the UI
            # We will fill in the archive list once the item is expanded
            self._friend_list.insert(parent, 0)


class YmArchive:
    def __init__(self, username, file_path):
        self._username = username.encode('utf-8')
        self._msg_list = []
        self._parse_file(file_path)

    @property
    def messages(self):
        return self._msg_list

    def _parse_file(self, file_path):
        with open(file_path, 'rb') as f:
            while True:
                timestamp_buf = f.read(5)
                if len(timestamp_buf) == 0:
                    break
                timestamp = struct.unpack('Ib', timestamp_buf)[0]
                YmArchive._ensure_3_zeros(f)
                is_received = f.read(1)[0] == 1
                YmArchive._ensure_3_zeros(f)
                ln = f.read(1)[0]
                YmArchive._ensure_3_zeros(f)
                msg = YmMsg(datetime.datetime.fromtimestamp(timestamp), is_received)
                if ln > 0:
                    msg.set_content(self._decrypt(f.read(ln)))
                f.read(4) # Skip the footer
                self._msg_list.append(msg)

    def _decrypt(self, raw_msg):
        ln = len(raw_msg)
        output = bytearray(ln)
        for i in range(ln):
            output[i] = raw_msg[i] ^ self._username[i % len(self._username)]
        return output.decode('utf-8')


    @staticmethod
    def _ensure_3_zeros(f):
        pos = f.tell()
        buf = f.read(3)
        for b in buf:
            if b != 0:
                raise Exception('3 all-zero bytes expected at %u' % pos)


class YmMsg:
    font_regex = re.compile(r'<font.*>')
    color_regex = re.compile(r'\x1B\[#[0-9A-Fa-f]{6}m')
    def __init__(self, timestamp, is_received):
        self._timestamp = timestamp
        self._is_received = is_received
        self._content = ''

    def set_content(self, new_content):
        # Strip the format specifier for now
        # In future we may want to parse and process those formats
        self._content = YmMsg._strip_format(new_content)

    @property
    def timestamp(self):
        return self._timestamp

    @property
    def is_received(self):
        return self._is_received

    @property
    def content(self):
        return self._content

    @staticmethod
    def _strip_format(input):
        result = input
        result = YmMsg.font_regex.sub('', result)
        result = YmMsg.color_regex.sub('', result)
        return result


if __name__ == '__main__':
    main()