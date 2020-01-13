# Supported styles

## Supported text styles and alignments

Note: You must import `NetOdt.Enumerations` to use text styles.

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

## Supported styles-sheets

Note: You must import `NetOdt.Enumerations` to use style-sheets.

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

## Supported colors

Note:
* You must import `System.Drawing` to use colors
* Alpha-channel is supported and will ignored

Use a color constant like `Color.Red`

or take it from RGB code like `Color.FromArgb(128, 128, 128)`
