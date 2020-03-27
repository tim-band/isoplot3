# Isoplot3 OOo

Isoplot3 does not work on any modern system. This project is an attempt to make
it work on OpenOffice/LibreOffice.

Isoplot is a collection of Visual Basic macros for Excel.

## Work to be done

* ~~Get it compiling~~ Done
* Get Main.Isoplot function running so that all the buttons appear
* Get each example working

## Differences between Visual Basic and OpenOffice Basic Syntax

OpenOffice uses a variant of Basic that is largely compatible with Visual Basic.

The following are issues that had to be worked around:

### lack of support for Next j, i

OpenOffice equivalent:

```vb
Next j: Next i
```

### Different scoping of line labels

Visual Basic line levels are local to a function, whereas OpenOffice Basic
labels are local to a module. Therefore duplicate labels must be deduplicated.

### ReDim not supported for Type members

For some reason, OpenOffice Basic does not support the following syntax:

```vb
ReDim MyType.member(NewSize)
```

A workaround is:

```vb
Dim MyType_member(NewSize)
MyType.member = MyType.member
```

### Only one member per line

```vb
Type MyType
  X As Double: Y As Double
End Type
```

is not permitted in OpenOffice Basic, each member must be on its own line.

### Type signifiers are significant

`name$` and `name` seem to refer to the same object in Visual Basic, but are different in
OpenOffice Basic. They must therefore all be unified.
