# pptxtract

Command line utility to extract references to external linked and embedded files and URLs to websites. Works with PowerPoint 2007 or later. Requires just the JDK and PicoCLI, no other dependencies yet.

## Usage

```shell
pptxtract TestData\powerpoint\Slideshow.pptx
Slideshow.pptx;File1.txt
Slideshow.pptx;image.bmp
Slideshow.pptx;File3.ini
Slideshow.pptx;File2.csv
```

or:

```shell
type pptfiles.txt | pptxtract
Slideshow1.pptx;File1.txt
Slideshow2.pptx;image.bmp
Slideshow3.pptx;File3.ini
Slideshow4.pptx;File2.csv
```

## Command Line Help

```
Usage: pptxtract [-hoVx] FILE...
      FILE...                PowerPoint FILE(S) where paths to embedded or
                               linked documents shall be extracted. Must be of
                               PowerPoint 2007 format (or later versions).
                               Older *.ppt files must be converted into
                               PowerPoint 2007 (or newer) format before use.
  -h, --help                 Show this help message and exit.
  -o                         When extracting embedded files, this option will
                               force overwriting existing files.
  -V, --version              Print version information and exit.
  -x, --extract-embeddings   When set, embedded files such as *.docx, *.xlsx or
                               other *.pptx files will be extracted.
```
