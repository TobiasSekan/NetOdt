# Supported styles

## Supported text styles and alignments

* Bold
* Italic
* Underline
  * Simple
  * Double
  * Bold
  * Dotted
  * Dotted bold
  * Dash
  * Long-dash
  * Dot-dash
  * Dot-dot-dash
  * Wave
* Stroke
* Superscript
* Subscript

1. Left
2. Center
3. Right
4. Justify

For style combinations you must combine the text styling via a bitwise or `|`, like
```
TextStyle.Center | TextStyle.Bold

TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSimple
```

Note: You can't combine different underlines, alignments and scripts, like
```
TextStyle.Left | TextStyle.Right

TextStyle.UnderlineBold | TextStyle.UnderlineWave

TextStyle.Superscript | TextStyle.Subscript
```

## Supported styles sheets

You can only use one style sheet for a text passage.

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
