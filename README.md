# massprint.vbs - generate lots of print jobs

This Visual Basic Script can take a set of Word, Excel and Powerpoint documents and print them a specified number of times, using the respective Microsoft products on your machine. The purpose is to easily create a huge amount of print jobs using different types of documents.

## motivation

Colleagues of mine wanted to generate a huge number of print jobs to test the correct functioning of a new load balancer. Instead of opening e.g. Word and hitting the print button a few thousand times, they asked me to cook up a script that could do the job.
_(Of course no trees were harmed during our testing, the print queue was on hold the hole time and all the jobs were deleted.)_

## usage

The script is desigend to use a given set of documents that you need to place in subfolders as follows:

```text
massprint.vbs
├───excel
├───powerpoint
└───word
```

It only recognizes files with the extensions `.xlsx`, `.pptx` and `.docx`, so files with different extensions are ignored. This repository provides sample documents in the correct folder structure.

Open the Windows command prompt and run e.g.

```cmd
  cscript massprint.vbs 10
```

to print the whole set of documents 10 times. With one document per folder, this would generate 30 print jobs in total.

Please take note:

- Documents will be printed to the default printer.
- You can break a running script with `CTRL+C` any time.
- Word, Excel and Powerpoint will open visibly on your computer. This is on purpose: if you quit the script or an error occurs, you would be left with hidden open instances of Word, Excel and Powerpoint without the means to close them.

## License

[MIT](LICENSE) (c) 2020 Generali Informatik
