# Supported styles

## Supported text styles and alignments

* Bold
* Italic
* Underline (single stroke)
* Stroke
* Superscript
* Subscript

* Left
* Center
* Right
* Justify

For style combinations you must combine the text styling via a bitwise or `|`, like
```
TextStyle.Center | TextStyle.Bold

TextStyle.Bold | TextStyle.Italic | TextStyle.Underline
```

Note: You can't combine different alignments, like
```
TextStyle.Left | TextStyle.Right
```

Note: You can't combine superscript and subscript
```
TextStyle.Superscript | TextStyle.Subscript
```

## Supported header styles

You can only use one style for a header.
* Title
* Heading level 1
* Heading level 2
* Heading level 3
* Heading level 4
* Heading level 5
* Heading level 6
* Heading level 7
* Heading level 8
* Heading level 9
* Heading level 10
* Subtitle
* Signature
* Quotations
* Endnote
* Footnote
