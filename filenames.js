const Excel = require('exceljs');
const fs = require('fs');

//Set your images directory
const imageFolder = './Images'

//Set your excelify file & sheet
const excelifyFile = 'filename.xlsx'
const sheetName = 'Sheet 1'

//Enter column numbers
const SKUColumnNumber = 34;
const IMGUrlColumnNumber = 18;

//Empty array to push filenames into
const fileNames = [];

const workbook = new Excel.Workbook();

//get all file names from images folder, add to array
fs.readdirSync(imageFolder).forEach(file => {
    fileNames.push(file)
  });

console.log('Processing...')



// Read from the excelify doc
workbook.xlsx.readFile(excelifyFile)
  .then( () => {
    // Select correct sheet
    const worksheet = workbook.getWorksheet(sheetName);

    //Double check that these column numbers in getColumn correspond to your sheet
    const SKUColumn = worksheet.getColumn(SKUColumnNumber);
    const imgColumn = worksheet.getColumn(IMGUrlColumnNumber);

    const prefix = 'https://example.com/media/images';

    // For each cell/row in the SKU column
    SKUColumn.eachCell( (SKUCell) => {

        const serverFileNames = [];

        //Dont run for the whole sheet
        if (SKUCell.value) {
            
            // Loop through file names
            for (let i = 0; i < fileNames.length; i++) {

                // If filename is found in the cell
                if ( fileNames[i].includes(SKUCell.value) ) {

                    
                    // Create new file name with server url prefix and push to an array, semicolon separated, then transform array to string
                    const imgUrl = prefix + fileNames[i];
                    serverFileNames.push(imgUrl);
                    const fileNamesString = serverFileNames.join(';').toString()
                    
                    // Find the image column cell that's in the same row as the sku & enter the file array string
                    imgColumn.eachCell(function(imgCell) {
                        if (imgCell.row === SKUCell.row) {
                            imgCell.value = fileNamesString;
                        }
                    })
                    
                }
                
            }

        }


   
        
    })

         // Save back to file
    return workbook.xlsx.writeFile(excelifyFile)

  })
  .catch( error => {
      console.error('Error: ', error)
  });


  