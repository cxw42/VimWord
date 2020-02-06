# VimWord

Word VBA add-in permitting you to use a subset of normal-mode editing commands
in Word.

## Installation

 - Exit Word
 - Run `Install VimWord.bat`
   - Alternatively put `VimWord.dotm` in `...\Word\Startup` or load
     through `Add-Ins | Manage: Word Add-ins`.
 - Map a key to `VimDoCommand` (I use `Ctrl`+`;` because it's easy to type on
   my keyboard.)

### Mapping a key

- Right-click the Ribbon and select `Customize the Ribbon...`.
- At the bottom of the `Word Options` dialog that appears, you should
  see a `Keyboard Shortcuts: Customize` button.  Press it.
- In the dialog box that appears, under `Categories`, select `Macros`.
- Under `Macros`, select `VimDoCommand`.
- Click in the box under `Press new shortcut key:`.
- Press the key you want to map
- In the bottom-left, click `Assign`.
- In the bottom-right, click `Close`.
- Back in the `Word Options` dialog, press `OK`.

## Usage

 - Hit the key you mapped, then enter a normal-mode command (e.g.,
   `diw`).  Currently, the supported operators are `d` (delete), `y` (copy),
   and `v` (select).  `c` (change) is also supported but doesn't do anything
   other than select the text.

## License

Copyright (c) 2018--2020 Christopher White.
Portions Copyright (c) 2020 D3 Engineering, LLC.

Licensed CC-BY-NC-SA 4.0 or,
at your option, any later version.  For the avoidance of doubt, merely
using VimWord at work does not automatically make the use commercial.
