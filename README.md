## AlphaBlendImage Control

Poor Man's Transparent Image Control

### Description

`AlphaBlendImage` control is built-in `VB.Image` control replacement (sort of) that supports alpha transparent images through GDI+. Standard OLE automation `StdPicture` objects can load 32-bit alpha transparent images in `vbPicTypeIcon` subtype, although few controls paint the alpha channel on these `StdPicture`s. This `AlphaBlendImage` control brings this support.

### API

The control's public `GdipLoadPicture` method can load 32-bit alpha transparent PNGs in `StdPicture` objects which can be assigned to control's `Picture` property. In the sample `Form1` such alpha transparent image is loaded to a `StdPicture` and is assigned both to a built-in `VB.Image` control and to an `AlphaBlendImage` control to compare difference in output.

The control supports `Opacity` property for "global" control transparency level (in addition to per-pixel alpha). 

The control supports `MaskColor` property for color-key transparency (in addition to per-pixel alpha).

The control supports `Rotation` property which rotates the assigned image (in degrees). 

The control supports `Zoom` property which scales the image (only when `Stretch` is off).

The control is windowless and cannot get focus. Its `AutoRedraw` property controls if repaint is cached to memory 32-bit DIB for faster redraws.