# -*- coding: utf-8 -*-

import os
from .folder_path import folder_path


class Config(object):

    def __init__(self):
        self.__path = None
        self._config_data = {}

    @property
    def path(self):
        if self.__path is None:
            path = os.path.join(folder_path.AppData, '.pyWinVirtualDesktop')
            if not os.path.exists(path):
                os.mkdir(path)

            self.path = path

        return self.__path

    @path.setter
    def path(self, path):
        if self.__path is not None and self._config_data:
            self.save()

        self._config_data = {}

        if os.path.exists(path) and os.path.isdir(path):
            os.path.join(path, 'config.data')

        if os.path.exists(path):
            pyWinVirtualDesktop = __import__(__name__.split('.')[0])

            desktop_ids = pyWinVirtualDesktop.desktop_ids

            with open(path, 'r') as f:
                data = f.read().split('\n')

            for line in data:
                guid, name = line.split(':', 1)

                if guid in desktop_ids:
                    self._config_data[guid] = name

        self.__path = path

    def get_name(self, guid):
        if guid in self._config_data:
            return self._config_data[guid]

        pyWinVirtualDesktop = __import__(__name__.split('.')[0])

        desktop_ids = pyWinVirtualDesktop.desktop_ids

        if guid in desktop_ids:
            self._config_data[guid] = (
                'Desktop ' + str(desktop_ids.index(guid) + 1)
            )

            return self._config_data[guid]

    def set_name(self, guid, name):
        self._config_data[guid] = name

    def save(self):
        output = []

        for guid, name in self._config_data.items():
            output += [guid + ':' + name]

        with open(self.path, 'w') as f:
            f.write('\n'.join(output))


Config = Config()
