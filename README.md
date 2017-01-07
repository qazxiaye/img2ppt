This tool uses a batch of images as backgrounds of slides in a generated powerpoint.

The images should be in the working directory and in jpg/png/gif format.

The tool is based on project ['powerpoint'](https://github.com/pythonicrubyist/powerpoint). It cooperates well with [Smallpdf](https://smallpdf.com/cn/pdf-to-jpg), a downloaded zip file can be the parameter of the script, unzip will be done automatically.

Require:
  gem install powerpoint

Usage:
  ./img2ppt : generate ppt using all the images in current folder as backgrounds
or
  ./img2ppt xxx.zip : unzip the zip file, and generate ppt based on unziped images in current foler

License: WTFPL

