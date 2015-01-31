enchant-windows-provider
========================

This is an [Enchant](http://www.abisource.com/projects/enchant/) provider that
uses the Windows [Spell Checking API](https://msdn.microsoft.com/en-us/library/windows/desktop/hh869852%28v=vs.85%29.aspx)
(introduced in Windows 8) as a backend.

Enchant is a spell-checker abstraction layer that is used by a number of
open-source programs (such as HexChat and Pidgin). With the addition of this
provider, you can have all such programs using a common, system dictionary.

Note that this project does not provide Enchant itself. If you're interested
in precompiled binaries for Enchant, the [HexChat GTK+ Bundle](https://hexchat.github.io/gtk-win32/)
might be of use to you.

Installation
============

The enchant_windows.dll can be placed directly into the 'lib\enchant' directory
of whatever Enchant-using program you want to use this with. Examples:

- Pidgin: C:\Program Files (x86)\Pidgin\spellcheck\lib\enchant
- HexChat: C:\Program Files\HexChat\lib\enchant

Development
===========

The code was developed using Visual Studio 2013 Express and compiles with same.
It might be possible to get it to build with 2012 if you work around not having
std::make_unique.

License
=======

LGPL, because Enchant is. See LICENSE.
