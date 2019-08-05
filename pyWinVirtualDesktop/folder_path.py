# -*- coding: utf-8 -*-

import os

import ctypes
from ctypes import HRESULT
from ctypes.wintypes import DWORD, HANDLE, HWND, LPWSTR, INT


kernel32 = ctypes.windll.Kernel32
shell32 = ctypes.windll.Shell32

GetTempPathW = kernel32.GetTempPathW
GetTempPathW.restype = DWORD
GetTempPathW.argtypes = [DWORD, LPWSTR]

SHGetFolderPathW = shell32.SHGetFolderPathW
SHGetFolderPathW.restype = HRESULT
SHGetFolderPathW.argtypes = [HWND, INT, HANDLE, DWORD, LPWSTR]

SHGFP_TYPE_CURRENT = 0
MAX_PATH = 260
CSIDL_FLAG_DONT_VERIFY = 16384


# Internet Explorer (icon on desktop)
CSIDL_INTERNET = 0x0001
# Start Menu\Programs
CSIDL_PROGRAMS = 0x0002
# My Computer\Control Panel
CSIDL_CONTROLS = 0x0003
# My Computer\Printers
CSIDL_PRINTERS = 0x0004
# My Documents
CSIDL_PERSONAL = 0x0005
# <user name>\Favorites
CSIDL_FAVORITES = 0x0006
# Start Menu\Programs\Startup
CSIDL_STARTUP = 0x0007
# <user name>\Recent
CSIDL_RECENT = 0x0008
# <user name>\SendTo
CSIDL_SENDTO = 0x0009
# <desktop>\Recycle Bin
CSIDL_BITBUCKET = 0x000A
# <user name>\Start Menu
CSIDL_STARTMENU = 0x000B
# Personal was just a silly name for My Documents
CSIDL_MYDOCUMENTS = CSIDL_PERSONAL
# "My Music" folder
CSIDL_MYMUSIC = 0x000D
# "My Videos" folder
CSIDL_MYVIDEO = 0x000E
# <user name>\Desktop
CSIDL_DESKTOPDIRECTORY = 0x0010
# My Computer
CSIDL_DRIVES = 0x0011
# Network Neighborhood (My Network Places)
CSIDL_NETWORK = 0x0012
# <user name>
CSIDL_NETHOOD = 0x0013
# windowsonts
CSIDL_FONTS = 0x0014
CSIDL_TEMPLATES = 0x0015
# All Users\Start Menu
CSIDL_COMMON_STARTMENU = 0x0016
# All Users\Start Menu\Programs
CSIDL_COMMON_PROGRAMS = 0x0017
# All Users\Startup
CSIDL_COMMON_STARTUP = 0x0018
# All Users\Desktop
CSIDL_COMMON_DESKTOPDIRECTORY = 0x0019
# <user name>\Application Data
CSIDL_APPDATA = 0x001A
# <user name>\PrintHood
CSIDL_PRINTHOOD = 0x001B
# <user name>\Local Settings\Applicaiton Data (non roaming)
CSIDL_LOCAL_APPDATA = 0x001C
# non localized startup
CSIDL_ALTSTARTUP = 0x001D
# non localized common startup
CSIDL_COMMON_ALTSTARTUP = 0x001E
CSIDL_COMMON_FAVORITES = 0x001F
CSIDL_INTERNET_CACHE = 0x0020
CSIDL_COOKIES = 0x0021
CSIDL_HISTORY = 0x0022
# All Users\Application Data
CSIDL_COMMON_APPDATA = 0x0023
# GetWindowsDirectory()
CSIDL_WINDOWS = 0x0024
# GetSystemDirectory()
CSIDL_SYSTEM = 0x0025
# C:\Program Files
CSIDL_PROGRAM_FILES = 0x0026
# C:\Program Files\My Pictures
CSIDL_MYPICTURES = 0x0027
# USERPROFILE
CSIDL_PROFILE = 0x0028
# x86 system directory on RISC
CSIDL_SYSTEMX86 = 0x0029
# x86 C:\Program Files on RISC
CSIDL_PROGRAM_FILESX86 = 0x002A
# C:\Program Files\Common
CSIDL_PROGRAM_FILES_COMMON = 0x002B
# x86 Program Files\Common on RISC
CSIDL_PROGRAM_FILES_COMMONX86 = 0x002C
# All Users\Templates
CSIDL_COMMON_TEMPLATES = 0x002D
# All Users\Documents
CSIDL_COMMON_DOCUMENTS = 0x002E
# All Users\Start Menu\Programs\Administrative Tools
CSIDL_COMMON_ADMINTOOLS = 0x002F
# <user name>\Start Menu\Programs\Administrative Tools
CSIDL_ADMINTOOLS = 0x0030
# Network and Dial-up Connections
CSIDL_CONNECTIONS = 0x0031
# All Users\My Music
CSIDL_COMMON_MUSIC = 0x0035
# All Users\My Pictures
CSIDL_COMMON_PICTURES = 0x0036
# All Users\My Video
CSIDL_COMMON_VIDEO = 0x0037
# Resource Direcotry
CSIDL_RESOURCES = 0x0038
# Localized Resource Direcotry
CSIDL_RESOURCES_LOCALIZED = 0x0039
# Links to All Users OEM specific apps
CSIDL_COMMON_OEM_LINKS = 0x003A
# USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
CSIDL_CDBURN_AREA = 0x003B
# Computers Near Me (computered from Workgroup membership)
CSIDL_COMPUTERSNEARME = 0x003D


class FolderPath(object):

    def __get_path(self, csidl):
        buf = ctypes.create_unicode_buffer(MAX_PATH)

        SHGetFolderPathW(
            0,
            csidl | CSIDL_FLAG_DONT_VERIFY,
            0,
            SHGFP_TYPE_CURRENT,
            buf
        )
        path = buf.value

        return path

    @property
    def TemporaryFiles(self):
        """
        temporary directory
        """
        buf = ctypes.create_unicode_buffer(MAX_PATH)
        GetTempPathW(MAX_PATH, buf)
        return buf.value[:-1]

    @property
    def InternetExplorer(self):
        """
        Internet Explorer (icon on desktop)
        """
        return self.__get_path(CSIDL_INTERNET)

    @property
    def Programs(self):
        """
        Start Menu\Programs
        """
        return self.__get_path(CSIDL_PROGRAMS)

    @property
    def ControlPanel(self):
        """
        My Computer\Control Panel
        """
        return self.__get_path(CSIDL_CONTROLS)

    @property
    def Printers(self):
        """
        My Computer\Printers
        """
        return self.__get_path(CSIDL_PRINTERS)

    @property
    def Personal(self):
        """
        My Documents
        """
        return self.__get_path(CSIDL_PERSONAL)

    @property
    def Favorites(self):
        """
        <user name>\Favorites
        """
        return self.__get_path(CSIDL_FAVORITES)

    @property
    def Startup(self):
        """
        Start Menu\Programs\Startup
        """
        return self.__get_path(CSIDL_STARTUP)

    @property
    def Recent(self):
        """
        <user name>\Recent
        """
        return self.__get_path(CSIDL_RECENT)

    @property
    def SendTo(self):
        """
        <user name>\SendTo
        """
        return self.__get_path(CSIDL_SENDTO)

    @property
    def RecycleBin(self):
        """
        <desktop>\Recycle Bin
        """
        return self.__get_path(CSIDL_BITBUCKET)

    @property
    def StartMenu(self):
        """
        <user name>\Start Menu
        """
        return self.__get_path(CSIDL_STARTMENU)

    @property
    def MyDocuments(self):
        """
        Personal was just a silly name for My Documents
        """
        return self.__get_path(CSIDL_MYDOCUMENTS)

    @property
    def MyMusic(self):
        """
        "My Music" folder
        """
        return self.__get_path(CSIDL_MYMUSIC)

    @property
    def MyVideos(self):
        """
        "My Videos" folder
        """
        return self.__get_path(CSIDL_MYVIDEO)

    @property
    def Desktop(self):
        """
        <user name>\Desktop
        """
        return self.__get_path(CSIDL_DESKTOPDIRECTORY)

    @property
    def MyComputer(self):
        """
        My Computer
        """
        return self.__get_path(CSIDL_DRIVES)

    @property
    def MyNetworkPlaces(self):
        """
        Network Neighborhood (My Network Places)
        """
        return self.__get_path(CSIDL_NETWORK)

    @property
    def NetHood(self):
        """
        nethood
        """
        return self.__get_path(CSIDL_NETHOOD)

    @property
    def Fonts(self):
        """
        windows fonts
        """
        return self.__get_path(CSIDL_FONTS)

    @property
    def Templates(self):
        """
        Templates
        """
        return self.__get_path(CSIDL_TEMPLATES)

    @property
    def CommonStartMenu(self):
        """
        All Users\Start Menu
        """
        return self.__get_path(CSIDL_COMMON_STARTMENU)

    @property
    def CommonPrograms(self):
        """
        All Users\Start Menu\Programs
        """
        return self.__get_path(CSIDL_COMMON_PROGRAMS)

    @property
    def CommonStartup(self):
        """
        All Users\Startup
        """
        return self.__get_path(CSIDL_COMMON_STARTUP)

    @property
    def CommonDesktop(self):
        """
        All Users\Desktop
        """
        return self.__get_path(CSIDL_COMMON_DESKTOPDIRECTORY)

    @property
    def AppData(self):
        """
        <user name>\Application Data (roaming)
        """
        return self.__get_path(CSIDL_APPDATA)

    @property
    def PrintHood(self):
        """
        <user name>\PrintHood
        """
        return self.__get_path(CSIDL_PRINTHOOD)

    @property
    def LocalAppData(self):
        """
        <user name>\Local Settings\Applicaiton Data (non roaming)
        """
        return self.__get_path(CSIDL_LOCAL_APPDATA)

    @property
    def AltStartup(self):
        """
        non localized startup
        """
        return self.__get_path(CSIDL_ALTSTARTUP)

    @property
    def CommonAltStartup(self):
        """
        non localized common startup
        """
        return self.__get_path(CSIDL_COMMON_ALTSTARTUP)

    @property
    def CommonFavorites(self):
        """
        common favorites
        """
        return self.__get_path(CSIDL_COMMON_FAVORITES)

    @property
    def InternetCache(self):
        """
        internet cache
        """
        return self.__get_path(CSIDL_INTERNET_CACHE)

    @property
    def Cookies(self):
        """
        cookies
        """
        return self.__get_path(CSIDL_COOKIES)

    @property
    def History(self):
        """
        history
        """
        return self.__get_path(CSIDL_HISTORY)

    @property
    def CommonAppData(self):
        """
        All Users\Application Data
        """
        return self.__get_path(CSIDL_COMMON_APPDATA)

    @property
    def Windows(self):
        """
        GetWindowsDirectory()
        """
        return self.__get_path(CSIDL_WINDOWS)

    @property
    def System(self):
        """
        GetSystemDirectory()
        """
        return self.__get_path(CSIDL_SYSTEM)

    @property
    def ProgramFiles(self):
        """
        C:\Program Files
        """
        return self.__get_path(CSIDL_PROGRAM_FILES)

    @property
    def MyPictures(self):
        """
        C:\Program Files\My Pictures
        """
        return self.__get_path(CSIDL_MYPICTURES)

    @property
    def Profile(self):
        """
        USERPROFILE
        """
        return self.__get_path(CSIDL_PROFILE)

    @property
    def ProgramFilesX86(self):
        """
        C:\Program Files (x86)
        """
        return self.__get_path(CSIDL_PROGRAM_FILESX86)

    @property
    def ProgramFilesCommon(self):
        """
        C:\Program Files\Common
        """
        return self.__get_path(CSIDL_PROGRAM_FILES_COMMON)

    @property
    def CommonTemplates(self):
        """
        All Users\Templates
        """
        return self.__get_path(CSIDL_COMMON_TEMPLATES)

    @property
    def CommonDocuments(self):
        """
        All Users\Documents
        """
        return self.__get_path(CSIDL_COMMON_DOCUMENTS)

    @property
    def CommonAdminTools(self):
        """
        All Users\Start Menu\Programs\Administrative Tools
        """
        return self.__get_path(CSIDL_COMMON_ADMINTOOLS)

    @property
    def AdminTools(self):
        """
        <user name>\Start Menu\Programs\Administrative Tools
        """
        return self.__get_path(CSIDL_ADMINTOOLS)

    @property
    def Connections(self):
        """
        Network and Dial-up Connections
        """
        return self.__get_path(CSIDL_CONNECTIONS)

    @property
    def CommonMusic(self):
        """
        All Users\My Music
        """
        return self.__get_path(CSIDL_COMMON_MUSIC)

    @property
    def CommonPictures(self):
        """
        All Users\My Pictures
        """
        return self.__get_path(CSIDL_COMMON_PICTURES)

    @property
    def CommonVideos(self):
        """
        All Users\My Video
        """
        return self.__get_path(CSIDL_COMMON_VIDEO)

    @property
    def Resources(self):
        """
        Resource Directory
        """
        return self.__get_path(CSIDL_RESOURCES)

    @property
    def ResourcesLocalized(self):
        """
        Localized Resource Directory
        """
        return self.__get_path(CSIDL_RESOURCES_LOCALIZED)

    @property
    def CommonOEMLinks(self):
        """
        Links to All Users OEM specific apps
        """
        return self.__get_path(CSIDL_COMMON_OEM_LINKS)

    @property
    def CDBurnArea(self):
        """
        USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
        """
        return self.__get_path(CSIDL_CDBURN_AREA)

    @property
    def ComputersNearMe(self):
        """
        Computers Near Me (computers from Workgroup membership)
        """
        return self.__get_path(CSIDL_COMPUTERSNEARME)


folder_path = FolderPath()
