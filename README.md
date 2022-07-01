# Info
Historic Vault of Programs and code that interface with AOL and AIM

Let's repopulate the Gibson!

Please contribute your files to this repository!

Thank you to:
* Len from Lens Hell
* https://web.archive.org/web/20220321112058/http://kadeklizem.com/AOL%20Progs%20ARCHIVE.rar
* http://www.aciddr0p.net/
* https://koin.org
* https://progs.rexflex.net/
* https://github.com/darcfx/darcfx-submissions
* https://github.com/raysuelzer/ProgzRescue

## commiting large amount of files

FYI if you ever try to commit a lot of zip files you will probably run in to errors.

A way around that is to use the included file:
```
gcommitfile.sh
```

### Example Use
Recursively commiting all files in current directory:

```
find . -exec gcommitfile.sh {} \;
```

Recursively commiting all files in current directory but omiting directory "unsorted-zip":

```
find . -not -path "./unsorted-zip/*" -exec gcommitfile.sh {} \;
```


# Directory Details

## oldscool_windows_tools
Tools compatible with Windows XP (many later versions are not)

* 7zip - 7z2107.exe (decompression)
* autoruns.exe - See what starts with system (Sysinternals tool)
* TweakUiPowertoySetup.exe - *Awesome tool to tweak the GUI of Windows XP*
* ProcessExplorerNT.zip - task manager on steroids (Sysinternal tool)
* Notepad ++ - npp.7.9.2.Installer.exe
* WinCDEmu-4.1.exe - *Use this to mount ISO or IMGs*
* winhex.zip - Hex editor

# Other
Items unrelated to programming/proggies or windows tools

* nfos.zip - Old warez scene nfo files

## programming

Mostly Visual Baisc files for interacting with AOL and AIM

## programs
Compiled AOL AND AIM programs used to interact with AOL.  Also known as Proggies.

