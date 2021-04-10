# DeleteFilesOlderThan

DeleteFilesOlderThan deletes files matching the passed file pattern and which are older than the given time specification from a folder.

---

## Usage

`DeleteFilesOlderThan /time=<time specification> /path=<folder to delete files from> [/filepattern=<files to delete>[;<files to delete>]] [/subfolders=0|1] [/filessmallerthan=|/filesgreaterthan=<file size>] [/recyclebin=0|1] [/readonly=0|1] [/processpriority=i|b] [/hideconsole=0|1]`

   or  

`DeleteFilesOlderThan /t=<time specification> /p=<folder to delete files from> [/f=<files to delete>[;<files to delete>]] [/s=0|1] [/fst=|/fgt=<file size>] [/rb=0|1] [/r=0|1] [/pp=i|b] [/hc=0|1]`

e.g.

- `DeleteFilesOlderThan /time=2d /path=D:\MyTarget`  
Delete _all_ files in folder `D:\MyTarget` which are older than two days.

- `DeleteFilesOlderThan /t=3w /p=C:\MyTarget\Data /f=*.txt /s=1`  
Delete all `*.txt` files in folder `C:\MyTarget\Data` _and all subfolders (`/s`)_ which are older than three weeks.


_Pressing_ __&lt;ESC&gt;__ _any time will exit the program execution._

## Parameters

- `/t` or `/time`  
Time specification _(see below)_
- `/p` or `/path`  
(Start) folder
- `/f` or `/filepattern`  
File pattern. If omitted, __all__ files are scanned _(equals `/f=*.*`)_.
- `/s` or `/subfolders`  
Recurse subfolders yes(1) or no(0). If omitted, only the folder passed via `/p` is scanned for matching files _(equals `/s=0`)_.
- `/rb` or `/recyclebin`  
Delete to recycle bin instead of permanently delete. If omitted, defaults to 0 = delete files permanently.
- `/r` or `/readonly`  
Delete readonly files? If omitted, defaults to 0 = don't delete readonly files.
- `/pp` or `/processpriority`  
Set this process' priority to _Idle_ ('i' = lowest possible) or _Below normal_ ('b') in order to consume less _(mainly CPU)_ resources.
- `/hc` or `/hideconsole`  
Hide the application's (console) window? Yes(1) or no(0). Defaults to no.
- `/fst` or `/filessmallerthan`  
Only delete files _smaller_ than the specified file size _(see below how to pass file sizes)_.
- `/fgt` or `/filesgreaterthan`  
Only delete files _greater_ than the specified file size _(see below how to pass file sizes)_.

Please note that you may only use __either__ `/fst` __or__ `/fgt`. You can't use both parameters. If you happen to pass both parameters, the last one 'wins'.

You may specify more than one file pattern for the parameter `/f` by using ; _(semicolon)_ as a separator, e.g.  
`/f=*.doc;*.rtf` = deletes all `*.doc` and all `*.rtf` files from the specified folder.  
`/f=Backup*.bak;Log*.trn` = deletes all `Backup*.bak` and all `Log*.trn` files from the specified folder.

### Allowed time specification units for parameter /t are

    d = day   e.g. 1d
    w = week  e.g. 2w
    m = month e.g. 3m
    y = year  e.g. 4y

### Allowed file size units

    none = Byte, e.g. 100
    kb = Kilobyte, e.g. 100kb
    mb = Megabyte, e.g. 100mb
    gb = Gigabyte, e.g. 100gb
    tb = Terabyte, e.g. 100tb

_Please note_: 1 KB = 1024 byte, 1 MB = 1024 KB etc.

## Creating a log file

DeleteFilesOlderThan writes all output to STDOUT. So in order to produce a log file of its actions, simply redirect the output to a file via _'> log\_file\_name'_.
