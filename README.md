# excel-js-filenames
Helper to write image src filename arrays to excelify sheet


## Getting Started


Clone/download repo to your machine

NPM install (only installs exceljs so you could just do that if you want)

Edit all the variables in the file to define your image folder location, sheet name, column numbers etc

### To Run

```
node filenames.js
```


## FYI

This uses two libraries: [ExcelJS](https://github.com/exceljs/exceljs) to read & write excel files, and fs (comes with node by default) to read file names.

## Why

If you can be bothered reading, details below.

I was provided a google drive folder of images, which were to be uploaded to our server and the new url/s added to the Excelify import sheet for import. 

We asked the client to name their images using the SKU and append the image number to the end, eg: HDK-200-1.jpg. 

I then needed to append the server url to this filename eg: example.com/media/images/HDK-200-1.jpg and add it to the 'Image SRC' column in the Excelify import spreadsheet.

Most products had multiple images, however the image filenames didnt always progress in numeric sequence. Some products would have images 1, 7, 11 & 12 which would result in HDK-200-1.jpg, HDK-200-7.jpg etc. So having a concatenate formula in excel wouldnt cut it as i'd need to manually check the files to see how many & what image numbers applied to each SKU.

What i needed to do was for each product in the Excelify sheet, get it's SKU, then check the image folder for any image with that SKU, and get it's filename, then prepend the server url to that.


Then i needed to combine all that products new image urls into one string and separate them by semicolon, and then insert that into the Image SRC column.

That's what this does.

