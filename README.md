# vibreoffice

vibreoffice is an extension for Libreoffice and OpenOffice that brings some of
your favorite key bindings from vi/vim to your favorite office suite. It is
obviously not meant to be feature-complete, but hopefully will be useful to
both vi/vim neophytes and experts alike.

### Installation/Usage

The easiest way to install is to download the
[latest extension file](https://raw.github.com/seanyeh/vibreoffice/master/dist/vibreoffice-0.1.1.oxt)
and open it with LibreOffice/OpenOffice.

To enable/disable vibreoffice, simply select Tools -> Add-Ons -> vibreoffice.

If you really want to, you can build the .oxt file yourself by running
```shell
# replace 0.0.0 with your desired version number
VIBREOFFICE_VERSION="0.0.0" make extension
```
This will simply build the extension file from the template files in
`extension/template`. These template files were auto-generated using
[Extension Compiler](https://wiki.openoffice.org/wiki/Extensions_Packager#Download).


### Features

vibreoffice currently supports:
- Insert (`i`, `I`, `a`, `A`, `o`, `O`), Visual (`v`), Normal modes
- Movement keys: `hjkl`, `w`, `W`, `b`, `B`, `e`, `$`, `^`, `{}`, `()`, `C-d`, `C-u`
    - Search movement: `f`, `F`, `t`, `T`
- Number modifiers: e.g. `5w`, `4fa`
- Replace: `r`
- Deletion: `x`, `d`, `c`, `s`, `D`, `C`, `S`, `dd`, `cc`
    - Plus movement and number modifiers: e.g. `5dw`, `c3j`, `2dfe`
    - Delete a/inner block: e.g. `di(`, `da{`, `ci[`, `ci"`, `ca'`, `dit`
- Undo/redo: `u`, `C-r`

### Known differences/issues

If you are familiar with vi/vim, then vibreoffice should give very few
surprises. However, there are some differences, primarily due to word
processor-text editor differences or limitations of the LibreOffice API and/or
my patience.
- Currently, I am using LibreOffice's built-in word/sentence movement which
  differs from vi's. It's sort of broken now but I plan to fix it eventually.
- The concept of lines in a text editor is not quite analogous to that of a
  word processor. I made my best effort to incorporate the line analogy while keeping
  the spirit of word processing.
    - Unlike vi/vim, movement keys will wrap to the next line
    - Due to line wrapping, you may find your cursor move up/down a line for
      commands that would otherwise leave you in the same position (such as `dd`)
- vibreoffice does not have contextual awareness. What I mean by that is that
  it does not keep track of which parentheses/braces match. Hence, you may have
  unexpected behavior (using commands such as `di(`) if your document has
  syntatically uneven parentheses/braces or nesting of such symbols. I don't
  intend to fix this for now, as I don't believe this is a critical feature for
  word processing.
- Using `d`, `c` (or any of their variants) will temporarily bring you into
  Visual mode. This is intentional and should not have any noticeable effects.

vibreoffice is new, so it is bound to have plenty of bugs. Please let me know
if you run into anything!


### License
vibreoffice is released under the MIT License.
