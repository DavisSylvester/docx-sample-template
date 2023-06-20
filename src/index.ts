import fs from 'fs';   
import path from 'path';
import fsp from 'fs/promises';


import { Document, Packer, Paragraph, Table, TableRow, TextRun, TableCell, WidthType, BorderStyle } from 'docx';


const doc = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                        new TextRun({
                            text: "\tGithub is the best",
                            bold: true,
                        }),
                    ],
                }),
                new Table({
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: {
                                        size: 25,
                                        type: WidthType.PERCENTAGE,                                        
                                    },
                                    borders: {
                                        top: {                                            
                                            size: 1,
                                            style: BorderStyle.DASH_DOT_STROKED,
                                        },
                                    }, 
                                                                       
                                    children: [new Paragraph("hello")],
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});

// Used to export the file into a .docx file
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});