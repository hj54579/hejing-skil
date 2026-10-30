const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel
} = require("docx");

const CJK = "Microsoft YaHei";
const font = { ascii: "Arial", hAnsi: "Arial", eastAsia: CJK };

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font, size: 24 },
        paragraph: { spacing: { after: 160, line: 380 } }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font, color: "1F3864" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 0, keepNext: false, keepLines: false }
      }
    ]
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 240 },
          children: [new TextRun({ text: "入团一对一谈话内容", bold: true, size: 36, font, color: "1F3864" })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("一、入团动机")] }),
        new Paragraph({ children: [new TextRun("共青团是党领导的先进青年群团组织，是党的助手和后备军。受身边党团员长辈和优秀同学的影响，我一直渴望加入共青团，在先进组织中锤炼自己、提高觉悟，争取为集体和社会多做贡献，成为新时代有担当的青年。")] }),

        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("二、团基础知识")] }),
        new Paragraph({ children: [new TextRun("我知道共青团是党领导下的先进青年群团组织；其根本任务是培养社会主义建设者和接班人。我学习了团员应履行的义务和享有的权利，熟悉团旗、团徽、团歌及入团誓词和入团程序，认识到团员不仅是荣誉，更是一份责任。")] }),

        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("三、思想感悟")] }),
        new Paragraph({ children: [new TextRun("通过团课学习和实践，我更坚定了理想信念，懂得了把个人理想融入国家事业，也体会到集体主义和奉献精神的意义，学会严于律己、在反思中不断进步。")] }),

        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("四、学习生活表现")] }),
        new Paragraph({ children: [new TextRun("学习上勤奋刻苦、成绩较好，积极参加课外阅读和社会实践；生活上尊敬师长、团结同学、乐于助人，遵规守纪；集体活动中热心班级事务，积极参加志愿服务和文体活动，有较强的集体荣誉感。")] }),

        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("五、自我评价与今后努力方向")] }),
        new Paragraph({ children: [new TextRun("我积极向上、品行端正、乐于奉献，但理论学习不够深入、自我要求有时不够严格。今后我将加强政治理论学习，严格以团员标准要求自己，主动服务同学和集体，全面提升素质，争取早日成为一名合格的共青团员。")] }),

        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 320 }, children: [new TextRun({ text: "谈话人：　　　　　　", font })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "日期：　　年　　月　　日", font })] })
      ]
    }
  ]
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/workspace/tuan_talk/入团一对一谈话内容.docx", buffer);
  console.log("OK: 入团一对一谈话内容.docx");
});