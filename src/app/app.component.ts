import { Component } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { saveAs } from 'file-saver/FileSaver';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Media,
  PictureRun,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  VerticalPositionAlign,
  WidthType,
  Header,
  Footer,
  Table,
  TableRow,
  TableCell,
  HeightRule,
  Styles,
} from 'docx';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  constructor(private http: HttpClient) {}

  public download(): void {
    this.http
      .get(
        'https://raw.githubusercontent.com/dolanmiu/docx/ccd655ef8be3828f2c4b1feb3517a905f98409d9/demo/images/cat.jpg',
        { responseType: 'blob' }
      )
      .subscribe((val) => {
        const doc = new Document();

        new Response(val).arrayBuffer().then((buffer) => {
          const image = Media.addImage(doc, buffer, 120, 120, {
            floating: {
              horizontalPosition: {
                align: HorizontalPositionAlign.RIGHT,
                relative: HorizontalPositionRelativeFrom.COLUMN,
              },
              verticalPosition: {
                align: VerticalPositionAlign.TOP,
                relative: VerticalPositionRelativeFrom.LINE,
              },
            },
          });

          doc.Styles.createParagraphStyle('para', 'Para').font('Calibri');

          doc.addSection({
            headers: {
              default: new Header({
                children: [new Paragraph({ text: 'Header', style: 'para' })],
              }),
            },
            footers: {
              default: new Footer({
                children: [new Paragraph({ text: 'Footer', style: 'para' })],
              }),
            },
            children: [
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({ text: 'Hello world', style: 'para' }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({ text: 'Hello world', style: 'para' }),
                        ],
                      }),
                    ],
                    height: { height: 2000, rule: HeightRule.EXACT },
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({ text: 'Hello world', style: 'para' }),
                        ],
                      }),
                      new TableCell({
                        children: [new Paragraph('hello')],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          });

          // const table = doc
          //   .createTable(2, 2)
          //   .setWidth(WidthType.PERCENTAGE, "100%");

          // table
          //   .getCell(0, 0)
          //   .addContent(
          //     new Paragraph()
          //       .spacing({ before: 2000 })
          //       .heading3()
          //       .addRun(new TextRun("Strictly Private & Confidential").bold())
          //   )
          //   .addContent(
          //     new Paragraph().heading3().addRun(new TextRun("Address").bold())
          //   )
          //   .CellProperties.setWidth("70%", WidthType.PERCENTAGE);

          // table
          //   .getCell(0, 1)
          //   .addContent(
          //     new Paragraph()
          //       .spacing({ before: 2000 })
          //       .heading3()
          //       .addRun(new TextRun("30 September 2019"))
          //   )
          //   .CellProperties.setWidth("30%", WidthType.PERCENTAGE);

          // table
          //   .getCell(1, 0)
          //   .addContent(
          //     new Paragraph()
          //       .spacing({ before: 2000 })
          //       .heading2()
          //       .addRun(new TextRun("Title").bold())
          //   )
          //   .addContent(
          //     new Paragraph()
          //       .spacing({ after: 1000 })
          //       .addRun(new TextRun("Some text here..").bold())
          //   )
          //   .addContent(
          //     new Paragraph()
          //       .heading2()
          //       .addRun(new TextRun("What to do").bold())
          //   )
          //   .addContent(
          //     new Paragraph()
          //       .spacing({ after: 1000 })
          //       .addRun(new TextRun("Some text here..").bold())
          //   )
          //   .addContent(
          //     new Paragraph().addRun(
          //       new TextRun(
          //         "If you have any questions, please contact our Customer Care Team."
          //       )
          //     )
          //   )
          //  .addContent(
          //   new Paragraph()
          //     .spacing({ before: 200 })
          //     .addRun(new TextRun("[Signature]"))
          // )
          // .addContent(
          //   new Paragraph()
          //     .spacing({ before: 200 })
          //     .addRun(new TextRun("Richard Basham-Jones"))
          // )
          // .addContent(
          //   new Paragraph()
          //     .spacing({ before: 200 })
          //     .addRun(new TextRun("Head of Customer Experience").italic())
          // )
          // .CellProperties.setWidth("50%", WidthType.PERCENTAGE);

          Packer.toBlob(doc).then((blob) => {
            console.log(blob);
            saveAs(blob, 'example.docx');
            console.log('Document created successfully');
          });

          // doc.addSection({
          //   footers: {
          //     default: {
          //       children: [
          //         new Paragraph(
          //           "The Royal London Mutual Insurance Society Limited is authorised by the Prudential Regulation Authority and regulated by the Financial Conduct Authority and the Prudential Regulation Authority. The firm is on the Financial Services Register, registration number 117672. It provides life assurance and pensions. Registered in England and Wales number 99064. Registered office: 55 Gracechurch Street, London, EC3V 0RL. Royal London Marketing Limited is authorised and regulated by the Financial Conduct Authority and introduces Royal London's customers to other insurance companies. The firm is on the Financial Services Register, registration number 302391.  Registered in England and Wales number 4414137.  Registered office: 55 Gracechurch Street, London, EC3V 0RL."
          //         )
          //       ]
          //     })
          //   },
          //   children: [new Paragraph(image), new Paragraph("Hello World")],
          //   margins: { top: 1000 }
          // });

          // new Response(val).arrayBuffer().then(buffer => {
          //   const image = Media.addImage(doc, buffer, 120, 120, {
          //     floating: {
          //       horizontalPosition: {
          //         align: HorizontalPositionAlign.RIGHT,
          //         relative: HorizontalPositionRelativeFrom.COLUMN
          //       },
          //       verticalPosition: {
          //         align: VerticalPositionAlign.TOP,
          //         relative: VerticalPositionRelativeFrom.LINE
          //       }
          //     }
          //   });

          //   doc.addSection({
          //     // headers: {
          //     //   default: new Header({
          //     //     children: [new Paragraph(image)]
          //     //   })
          //     // },
          //     footers: {
          //       default: new Footer({
          //         children: [
          //           new Paragraph(
          //             "The Royal London Mutual Insurance Society Limited is authorised by the Prudential Regulation Authority and regulated by the Financial Conduct Authority and the Prudential Regulation Authority. The firm is on the Financial Services Register, registration number 117672. It provides life assurance and pensions. Registered in England and Wales number 99064. Registered office: 55 Gracechurch Street, London, EC3V 0RL. Royal London Marketing Limited is authorised and regulated by the Financial Conduct Authority and introduces Royal London's customers to other insurance companies. The firm is on the Financial Services Register, registration number 302391.  Registered in England and Wales number 4414137.  Registered office: 55 Gracechurch Street, London, EC3V 0RL."
          //           )
          //         ]
          //       })
          //     },
          //     children: [new Paragraph(image), new Paragraph("Hello World")],
          //     margins: { top: 1000 }
          //   });
        });
      });
  }
}
