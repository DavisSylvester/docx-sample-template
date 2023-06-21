import fs from 'fs';   
import path from 'path';
import fsp from 'fs/promises';


import { Document, Packer, Paragraph, Table, TableRow, TextRun, TableCell, WidthType, BorderStyle, ShadingType, IShadingAttributesProperties } from 'docx';

const b: IShadingAttributesProperties = {
    color: "#FFFFFF",
    fill: "#880808"
};

const v: IShadingAttributesProperties = {
    color: "#FFFFFF",
    fill: "#EA3B52"
};

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
                                    shading: b,                     
                                    children: [new Paragraph("hello")],
                                }),

                                new TableCell({
                                    width: {
                                        size: 75,
                                        type: WidthType.PERCENTAGE,                                        
                                    },
                                    borders: {                                        
                                        right: {                                            
                                            size: 1,
                                            style: BorderStyle.DASH_DOT_STROKED,
                                        },
                                    }, 
                                    shading: v,                     
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