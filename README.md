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

## Usage

 - Hit the key you mapped, then enter a normal-mode command (e.g.,
   `diw`).  Currently only the `d` and `y` operators are supported,
   plus `;` or `'` (not in Vim) to just go there.  For example, `;10l`
   will move 10 characters right.

## Limitations

 - Counts are only supported on `h` and `l`
 - Vertical motion is not yet supported
 - Numerous others! :)

## License

CC-BY-NC-SA 4.0 or, at your option, any later version.

