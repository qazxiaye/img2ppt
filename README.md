This tool uses a batch of input images as backgrounds of generates powerpoint.

Input images should be in the working directory and in jpg/png/gif format. The tool is based on project ['powerpoint'](https://github.com/pythonicrubyist/powerpoint). It cooperates well with [Smallpdf](https://smallpdf.com/cn/pdf-to-jpg), a zip file can also be input of the script.

<br />

Require:
 - `gem install powerpoint`

Usage:
 - `./img2ppt.rb` : generate ppt using images under working directory as backgrounds
 - `./img2ppt.rb xxx.zip` : unzip the input zip file, then generate ppt file based on unziped images

<br />

License: WTFPL

