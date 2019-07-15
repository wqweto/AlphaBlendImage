## AlphaBlendImage Control

Poor Man's Transparent Image Control

### Description

AlphaBlendImage control is built-in Image control replacement (sort of) that supports alpha transparent images through GDI+. Built-in ole automation `StdPicture` can load 32-bit alpha transparent images in `vbPicTypeIcon` sub-type, although few controls support such `StdPicture`s. This `AlphaBlendImage` control does.

### API

It also has a public `GdipLoadPicture` method so that one can load 32-bit PNGs in `StdPicture`s and assign these to control's `Picture` property. In the sample `Form1` such alpha transparent image is loaded to a `StdPicture` and is assigned both to a built-in `Image1` control and to a `AlphaBlendImage1` control to compare difference in output.

The control support `Opacity` property -- in addition to per pixel alpha this controls "global" control transparency. 

The control support `Rotation` property which rotates the assigned image. 

The control support `MaskColor` property for key-color transparency.

The control is windowless and cannot get focus. Its `AutoRedraw` property controls if repaint is cached to in-memory DIB for faster redraws.