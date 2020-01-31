# massprint.vbs - generate lots of print jobs

This Visual Basic Script can take a set of Word, Excel and Powerpoint documents and print them a specified number of times, using the respective Microsoft products on your machine. The purpose is to easily create a huge amount of print jobs using different types of documents.

## Motivation

Colleagues of mine wanted to generate a huge number of print jobs to test the correct functioning of a new load balancer. Instead of opening e.g. Word and hitting the print button a few thousand times, they asked me to cook up a script that could do the job.
_(Of course no trees were harmed during our testing, the print queue was on hold the hole time and all the jobs were deleted.)_

## Usage

The script has been developed on Windows 10 1803 and Office 2010. It'll likely run on older and newer versions as the utilized functionality is pretty basic. Pause your printer queue, open a command prompt and give it a spin:

```cmd
cscript massprint.vbs 1
```

This will print the provided set of sample documents one time, generating 3 print jobs in total.

You can use your own documents, as many and as diverse as you like, as long as you put them into the folder structure as below:

```text
massprint.vbs
├───excel
├───powerpoint
└───word
```

It only recognizes files with the extensions `.xlsx`, `.pptx` and `.docx`, so files with different extensions are ignored.

Please take note:

- Documents will be printed to the default printer.
- Documents will be printed in the following order: all Word documents > all Excel spreadsheets > all Powerpoint presentations.
- You can break the running script with `CTRL+C` at any time.
- Word, Excel and Powerpoint will open visibly on your computer. This is on purpose: if they were opened in a hidden state and you would break the script or an error occurs, you would be left with open instances of Word, Excel and Powerpoint without the means to close them.

## Status

This script has been successfully used to fulfill its purpose. There won't be any further development until the need for it comes up.

## License

[MIT](LICENSE) © 2020 Generali Deutschland Informatik Services GmbH
