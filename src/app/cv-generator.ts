import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, TabStopPosition, TabStopType, TextRun , Table ,TableRow, TableCell,WidthType, BorderStyle,VerticalAlign, Media} from "docx";
import * as fs from "fs";
import * as FileSaver from 'file-saver';
const PHONE_NUMBER = "07534563401";
const PROFILE_URL = "https://www.linkedin.com/in/dolan1";
const EMAIL = "docx@docx.com";

export class DocumentCreator {
    // tslint:disable-next-line: typedef
    public create([experiences, educations, skills, achivements]): Document {
        const document = new Document();
             const image = Media.addImage(
            document,
            FileSaver.readFileSync(
                'http://localhost:4200/assets/images/placeholder_profile.png',
            ),
        );  

        document.addSection({
            children: [
                new Paragraph({
                    text: "Dolan Miu",
                    heading: HeadingLevel.TITLE,
                }),
                this.createContactInfo(PHONE_NUMBER, PROFILE_URL, EMAIL),
                this.createHeading("Education"),
                ...educations
                    .map((education) => {
                        const arr: Paragraph[] = [];
                        arr.push(
                            this.createInstitutionHeader(education.schoolName, `${education.startDate.year} - ${education.endDate.year}`),
                        );
                        arr.push(this.createRoleText(`${education.fieldOfStudy} - ${education.degree}`));

                        const bulletPoints = this.splitParagraphIntoBullets(education.notes);
                        bulletPoints.forEach((bulletPoint) => {
                            arr.push(this.createBullet(bulletPoint));
                        });

                        return arr;
                    })
                    .reduce((prev, curr) => prev.concat(curr), []),
                this.createHeading("Experience"),
                ...experiences
                    .map((position) => {
                        const arr: Paragraph[] = [];

                        arr.push(
                            this.createInstitutionHeader(
                                position.company.name,
                                this.createPositionDateText(position.startDate, position.endDate, position.isCurrent),
                            ),
                        );
                        arr.push(this.createRoleText(position.title));

                        const bulletPoints = this.splitParagraphIntoBullets(position.summary);

                        bulletPoints.forEach((bulletPoint) => {
                            arr.push(this.createBullet(bulletPoint));
                        });

                        return arr;
                    })
                    .reduce((prev, curr) => prev.concat(curr), []),


                     ...experiences
                    .map((position) => {
                        const arr: Table[] = [];
                        
                        arr.push(
                          this.createTable(position.company.name, position.title, position.summary)
                        
                        );
                     

                        return arr;
                    })
                    .reduce((prev, curr) => prev.concat(curr), []),
                this.createHeading("Skills, Achievements and Interests"),
                this.createSubHeading("Skills"),
                this.createSkillList(skills),
                this.createSubHeading("Achievements"),
                ...this.createAchivementsList(achivements),
                this.createSubHeading("Interests"),
                this.createInterests("Programming, Technology, Music Production, Web Design, 3D Modelling, Dancing."),
                this.createHeading("References"),
                new Paragraph(
                    "Dr. Dean Mohamedally Director of Postgraduate Studies Department of Computer Science, University College London Malet Place, Bloomsbury, London WC1E d.mohamedally@ucl.ac.uk",
                ),
                new Paragraph("More references upon request"),
                new Paragraph({
                    text: "This CV was generated in real-time based on my Linked-In profile from my personal website www.dolan.bio.",
                    alignment: AlignmentType.CENTER,
                }),
            ],
        });

        return document;
    }

      public createTable(companyName, title, summary) {
     const table = new Table({
     //alignment: AlignmentType.CENTER,
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph(companyName),new Paragraph(title)],
                    margins: {
                        top: 100,
                        bottom: 100,
                        left: 100,
                        right: 100,
                    },
                    verticalAlign: VerticalAlign.CENTER,
                    
                }),
                 new TableCell({
                    children: [new Paragraph(summary)],
                    margins: {
                        top: 100,
                        bottom: 100,
                        left: 100,
                        right: 100,
                    },
                }),
            ],
        }),
 
    ],
 width: {
        size: 100,
        type: WidthType.AUTO,
    },
    columnWidths: [4000, 5000],
});


const table5 = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("0,0")],
                }),
                new TableCell({
                    children: [new Paragraph("0,1")],
                    rowSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("0,2")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("1,0")],
                }),
                new TableCell({
                    children: [new Paragraph("1,2")],
                    rowSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("2,0")],
                }),
                new TableCell({
                    children: [new Paragraph("2,1")],
                }),
            ],
        }),
    ],
    width: {
        size: 100,
        type: WidthType.PERCENTAGE,
    },
});

return table
    }

    public createContactInfo(phoneNumber: string, profileUrl: string, email: string): Paragraph {
        return new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new TextRun(`Mobile: ${phoneNumber} | LinkedIn: ${profileUrl} | Email: ${email}`),
                new TextRun("Address: 58 Elm Avenue, Kent ME4 6ER, UK").break(),
            ],
        });
    }

    public createHeading(text: string): Paragraph {
        return new Paragraph({
            text: text,
            heading: HeadingLevel.HEADING_1,
            thematicBreak: true,
        });
    }

    public createSubHeading(text: string): Paragraph {
        return new Paragraph({
            text: text,
            heading: HeadingLevel.HEADING_2,
        });
    }

    public createInstitutionHeader(institutionName: string, dateText: string): Paragraph {
        return new Paragraph({
            tabStops: [
                {
                    type: TabStopType.RIGHT,
                    position: TabStopPosition.MAX,
                },
            ],
            children: [
                new TextRun({
                    text: institutionName,
                    bold: true,
                }),
                new TextRun({
                    text: `\t${dateText}`,
                    bold: true,
                }),
            ],
        });
    }

    public createRoleText(roleText: string): Paragraph {
        return new Paragraph({
            children: [
                new TextRun({
                    text: roleText,
                    italics: true,
                }),
            ],
        });
    }

    public createBullet(text: string): Paragraph {
        return new Paragraph({
            text: text,
            bullet: {
                level: 0,
            },
        });
    }

    // tslint:disable-next-line:no-any
    public createSkillList(skills: any[]): Paragraph {
        return new Paragraph({
            children: [new TextRun(skills.map((skill) => skill.name).join(", ") + ".")],
        });
    }

    // tslint:disable-next-line:no-any
    public createAchivementsList(achivements: any[]): Paragraph[] {
        return achivements.map(
            (achievement) =>
                new Paragraph({
                    text: achievement.name,
                    bullet: {
                        level: 0,
                    },
                }),
        );
    }

    public createInterests(interests: string): Paragraph {
        return new Paragraph({
            children: [new TextRun(interests)],
        });
    }

    public splitParagraphIntoBullets(text: string): string[] {
        return text.split("\n\n");
    }

    // tslint:disable-next-line:no-any
    public createPositionDateText(startDate: any, endDate: any, isCurrent: boolean): string {
        const startDateText = this.getMonthFromInt(startDate.month) + ". " + startDate.year;
        const endDateText = isCurrent ? "Present" : `${this.getMonthFromInt(endDate.month)}. ${endDate.year}`;

        return `${startDateText} - ${endDateText}`;
    }

    public getMonthFromInt(value: number): string {
        switch (value) {
            case 1:
                return "Jan";
            case 2:
                return "Feb";
            case 3:
                return "Mar";
            case 4:
                return "Apr";
            case 5:
                return "May";
            case 6:
                return "Jun";
            case 7:
                return "Jul";
            case 8:
                return "Aug";
            case 9:
                return "Sept";
            case 10:
                return "Oct";
            case 11:
                return "Nov";
            case 12:
                return "Dec";
            default:
                return "N/A";
        }
    }
}