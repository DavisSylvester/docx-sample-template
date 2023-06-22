import fs from 'fs';   
import path from 'path';
import fsp from 'fs/promises';


import { Document, Packer, Paragraph, Table, TableRow, TextRun, TableCell, WidthType, BorderStyle, ShadingType, IShadingAttributesProperties, HeightRule, TextDirection, AlignmentType, HeadingLevel } from 'docx';

const b: IShadingAttributesProperties = {
    color: "#FFFFFF",
    fill: "#880808"
};

const v: IShadingAttributesProperties = {
    color: "#FFFFFF",
    fill: "#EA3B52"
};

const headerColumns = ["Column 1", "Column 2", "Column 3", "Column 4"];

const createTableHeader = (headers: string[]) => {
    
    const headerRow = headers.map((header) => {
        return new TableCell({
            width: {
                size: 100 / headers.length,
                type: WidthType.PERCENTAGE,                                        
            },
            borders: {
                top: {                                            
                    size: 1,
                    style: BorderStyle.SINGLE,
                },
                bottom: {                                            
                    size: 1,
                    style: BorderStyle.SINGLE,
                },
                left: {                                            
                    size: 1,
                    style: BorderStyle.SINGLE,
                },
                right: {                                            
                    size: 1,
                    style: BorderStyle.SINGLE,
                },
            }, 
            shading: b,
            children: [new Paragraph({
                // text: header,
                alignment: AlignmentType.CENTER,
                heading: HeadingLevel.HEADING_1,                
                children: [
                    new TextRun({
                        text: header,
                        font: 'Open Sans',
                        size: '14pt',
                        color: '#FFFFFF',
                        bold: true,
                        style: "Heading1",
                        
                    })
                ],
            })],
        });
    });
    
    
    const header = new TableRow({
        height: {
            value: `20pt`,
            rule: HeightRule.EXACT,
        },
        tableHeader: true,
        children: headerRow,
    });
    
    return header;
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
                        // new TableRow({
                        //     children: [
                        //         new TableCell({
                        //             width: {
                        //                 size: 25,
                        //                 type: WidthType.PERCENTAGE,                                        
                        //             },
                        //             borders: {
                        //                 top: {                                            
                        //                     size: 1,
                        //                     style: BorderStyle.DASH_DOT_STROKED,
                        //                 },
                        //             }, 
                        //             shading: b,                     
                        //             children: [new Paragraph("hello")],
                        //         }),

                        //         new TableCell({
                        //             width: {
                        //                 size: 75,
                        //                 type: WidthType.PERCENTAGE,                                        
                        //             },
                        //             borders: {                                        
                        //                 right: {                                            
                        //                     size: 1,
                        //                     style: BorderStyle.DASH_DOT_STROKED,
                        //                 },
                        //             }, 
                        //             shading: v,                     
                        //             children: [new Paragraph("hello")],
                        //         }),
                        //     ],
                        // }),
                        createTableHeader(headerColumns)
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