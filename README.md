# *Compass Imaging Router* DICOM Transfer Reporting

## PowerShell Script (v4.0+) to Parse DICOMs
Copies *largest* DICOM (.dcm) from source directory to new destination.
Then runs a tag dump to extract specific information about each study. 
Note: the largest file size is used to avoid inadvertently parsing presentation states (PR) or stuctured reports (SR).
Finally this script exports desired metadata to .csv/.xlsx, where each row = 1 instance of a DICOM transfer to the IMG_RTR. 


## Options: 
To exclude SWMC studies and toggle logging
```powershell
$PURGE_CSV_FILE= $False
$EXPORT_FIRST_DCM = $True
$DUMP_TAGS_XMLs = $True
$FILTER_PH_AETs = $False
$PURGE_ALL = $False
pattern:['ADAC_','AEGISWEB','FILA_','MEHC_','RSEND_','SWMC_','SW_','SW_CATH','VANC_']
```

## Directories:
```powershell
$dcmtk_path = "$pwd_parent_path\dcmtk-3.6.5-win32-dynamic\bin"
$src_path = "$pwd_parent_path\input\images"
$dst_path = "$pwd_parent_path\output\$base_filename$(Get-Month-Filename)"
IMG_RTR_MM-YYYY_DICOMs (for output dcm/tag dumps)
IMG_RTR_MM-YYYY_LOGs   (for Excel/text reports)
```

## Example Output:
![Screenshot](https://github.com/github-pdx/dicom.router.parser/blob/master/img/excel.export.png)
* [DICOM-based Excel Report](https://github.com/github-pdx/dicom.router.parser/blob/master/output/IMG_RTR_Transfers_06-09-19.xlsx)
* [DICOM Tag Dump](https://github.com/github-pdx/dicom.router.parser/blob/master/592dbbe4f8994a36abcf85b4d67e1abb.junk/9fe63f0a-d304-4a22-9e4b-f0ebe63f7f78.txt)
* [XML Tag Dump](https://github.com/github-pdx/dicom.router.parser/blob/master/592dbbe4f8994a36abcf85b4d67e1abb.junk/9fe63f0a-d304-4a22-9e4b-f0ebe63f7f78.xml)
* [Tag-based Excel Report](https://github.com/github-pdx/dicom.router.parser/blob/master/output/~dicom_tag_dumps.xlsx)


## *DICOM* Resources:
* [Compass Router](http://www.laurelbridge.com/pdf/Compass-User-Manual.pdf)
* [DMTk](https://dicom.offis.de/dcmtk.php.en)
* [dcm4che](https://dcm4che.atlassian.net/wiki/spaces/lib/overview)
* [Slicer3D](https://www.slicer.org/)
* [DICOM Cleaner](http://www.dclunie.com/pixelmed/software/webstart/DicomCleanerUsage.html)
* [MATLAB DICOM Toolbox](https://www.mathworks.com/help/images/scientific-file-formats.html)
* [SonicDICOM](https://sonicdicom.com/)
* [Weasis](https://nroduit.github.io/en/)
* [pydicom](https://pydicom.github.io/pydicom/stable/index.html)
* [DCF SDK](http://www.laurelbridge.com/products/dcf/)
* [DCF Examples](http://www.laurelbridge.com/docs/dcf34/ExampleDocs/)
* [GDCM](https://github.com/malaterre/GDCM)

## License:
[Licensed under the MIT License](LICENSE).


## **Dependencies:**
* [PowerShell v4.0+](https://www.microsoft.com/en-us/download/details.aspx?id=54616)
* [DCMTK 3.6.5-executable binaries](https://github.com/github-pdx/dicom.router.parser/tree/master/dcmtk-3.6.5-win32-dynamic)
* [Microsoft Visual C++ 2012 Redistributable Package (x64)](https://www.microsoft.com/en-us/download/details.aspx?id=30679)

When attempting to run \dcmtk-3.6.5-**win64**-dynamic\bin\dcmdump.exe to parse tag data, if you get the following error: 

![Screenshot](https://github.com/github-pdx/dicom.router.parser/blob/master/img/missing.MSVCP110.dll_(x64).png)

**Below is the fix:**

![Screenshot](https://github.com/github-pdx/dicom.router.parser/blob/master/img/install.C++.redistributable_(x64).png)

## Maintainer:
Emile Averill
