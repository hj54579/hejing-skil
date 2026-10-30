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
        run: { size: 32, bold: true, font, color: "1F3864" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 0, keepNext: false, keepLines: false }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font, color: "2E5395" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1, keepNext: false, keepLines: false }
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
        // 标题
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "入团一对一谈话内容", bold: true, size: 40, font, color: "1F3864" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 280 },
          children: [new TextRun({ text: "（入团动机 · 团基础知识 · 思想感悟 · 学习生活表现 · 自我评价与今后努力方向）", size: 20, font, color: "595959" })]
        }),

        // 一、入团动机
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("一、入团动机")] }),
        new Paragraph({ children: [new TextRun("我一直对共青团组织怀有真挚的崇敬与向往。共青团是中国共产党领导的先进青年的群团组织，是广大青年在实践中学习中国特色社会主义和共产主义思想的重要学校，是中国共产党的助手和后备军。加入共青团，是我长久以来的心愿，更是我人生成长道路上的一个明确目标。")] }),
        new Paragraph({ children: [new TextRun("首先，从家庭和身边榜样的影响来看，我的父母都是有着坚定政治信仰的人，家中的长辈和亲友里有不少共产党员和共青团员，他们平日里为人处世严谨踏实、乐于奉献，这让我从小就在心里埋下了一颗追求进步的种子。在学校里，我看到团员同学积极参加志愿服务、热心帮助他人、在学习上严格要求自己，他们身上那种积极向上的精神面貌深深感染着我，也让我更加渴望成为他们中的一员。")] }),
        new Paragraph({ children: [new TextRun("其次，从个人成长的追求来看，我希望能在一个先进的组织中得到锤炼和成长。加入共青团，意味着接受组织的教育和监督，意味着要对自己的言行提出更高的要求。我愿意以团员的标准鞭策自己，不断端正思想、提高觉悟、锤炼品格，在集体中发光发热。")] }),
        new Paragraph({ children: [new TextRun("再次，从时代责任的角度来看，作为新时代的青年，我深知自己肩负着民族复兴和国家建设的责任。加入共青团，与广大优秀青年一起投身于实现中华民族伟大复兴中国梦的伟大事业，是我义不容辞的选择。我希望通过自己的努力，为国家、为社会贡献一份青春力量。")] }),

        // 二、团基础知识
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("二、团基础知识")] }),
        new Paragraph({ children: [new TextRun("通过学习和培训，我对共青团的了解不断加深。中国共产主义青年团是中国共产党领导的先进青年的群团组织，是广大青年在实践中学习中国特色社会主义和共产主义思想的重要学校，是党的助手和后备军。团的基本知识我主要掌握了以下几个方面：")] }),
        new Paragraph({ children: [new TextRun({ text: "1. 团的章程与性质：", bold: true, font }), new TextRun("团的章程明确规定，团是党领导下的先进青年群团组织。团的建设必须坚持党的领导，坚持把培养社会主义建设者和接班人作为根本任务，把巩固和扩大党执政的青年群众基础作为政治责任。")] }),
        new Paragraph({ children: [new TextRun({ text: "2. 团的基本任务：", bold: true, font }), new TextRun("以习近平新时代中国特色社会主义思想为指导，团结带领广大青年，为实现中华民族伟大复兴的中国梦贡献智慧和力量。")] }),
        new Paragraph({ children: [new TextRun({ text: "3. 团员的义务与权利：", bold: true, font }), new TextRun("团员应当自觉履行学习的义务、宣传执行的义务、遵纪守法的义务、参加组织生活和完成组织交办任务的义务。同时，团员享有选举权、被选举权和表决权，以及参加团组织活动的权利、对团的工作提出建议和监督的权利等。")] }),
        new Paragraph({ children: [new TextRun({ text: "4. 团的标志：", bold: true, font }), new TextRun("团旗是团组织的象征，团徽是共青团员身份的标志，团歌《光荣啊，中国共青团》激励着一代又一代青年奋发向上。我熟悉团旗、团徽、团歌的基本含义和使用规范，也清楚入团誓词的内容和入团的基本程序。")] }),
        new Paragraph({ children: [new TextRun("通过系统学习，我认识到共青团员不仅是一种荣誉，更是一份沉甸甸的责任。只有真正在思想上、行动上向组织靠拢，才配得上“共青团员”这一光荣称号。")] }),

        // 三、思想感悟
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("三、思想感悟")] }),
        new Paragraph({ children: [new TextRun("在向团组织靠拢的过程中，我的思想认识经历了从浅显到逐步深刻的变化，收获颇丰。")] }),
        new Paragraph({ children: [new TextRun({ text: "第一，", bold: true, font }), new TextRun("我更加深刻地理解了信仰的力量。以前对“理想信念”的理解比较抽象，通过学习和参加团课、主题活动，我看到无数优秀青年为了国家和人民的事业甘于奉献、勇挑重担，我真正感受到，一个人只有把个人的理想追求融入到国家和民族的事业中去，才能拥有源源不断的动力。")] }),
        new Paragraph({ children: [new TextRun({ text: "第二，", bold: true, font }), new TextRun("我懂得了集体主义和奉献精神的意义。在参加志愿服务和班级集体活动的过程中，我体会到，团结协作、互帮互助能够汇聚成强大的力量。一个人的价值不在于索取，而在于为他人、为集体、为社会做出了多少贡献。这种感悟让我更加愿意主动服务同学、服务集体。")] }),
        new Paragraph({ children: [new TextRun({ text: "第三，", bold: true, font }), new TextRun("我认识到严于律己、不断反思的重要性。在对照团员标准查找差距的过程中，我发现自己在学习上偶尔有松懈、在生活中有时不善于主动承担等不足。正是有了这些反思，我才更加明确了自己的改进方向，也更加坚定了向团组织靠拢的决心。")] }),
        new Paragraph({ children: [new TextRun("总的来说，入团的过程是一个思想不断升华、自我不断完善的过程，它让我明白：进步不是一句口号，而是体现在每一天实实在在的行动之中。")] }),

        // 四、学习生活表现
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("四、学习生活表现")] }),
        new Paragraph({ children: [new TextRun({ text: "（一）学习方面。", bold: true, font }), new TextRun("我始终把学习作为第一要务。课堂上认真听讲、勤于思考、积极完成作业；课余时间主动阅读课外书籍，拓宽知识面。遇到难题不轻易放弃，愿意向老师和同学请教。我的学习成绩保持在班级前列，并多次在学科竞赛中获得名次。同时，我也不满足于课本知识，积极参加社会实践和研学活动，努力做到学以致用、知行合一。")] }),
        new Paragraph({ children: [new TextRun({ text: "（二）生活方面。", bold: true, font }), new TextRun("我注重养成良好的生活习惯，作息规律、勤俭节约、热爱劳动。在家尊敬长辈、主动分担家务；在学校尊敬老师、团结同学、乐于助人。当同学遇到困难时，我会主动伸出援手，尽己所能帮助他们。我严格遵守校规校纪，自觉维护班级和学校的良好秩序。")] }),
        new Paragraph({ children: [new TextRun({ text: "（三）集体活动方面。", bold: true, font }), new TextRun("我热心班级事务，积极参加学校组织的各类文体活动和志愿服务，比如运动会、文艺汇演、社区志愿服务、义务劳动等。在集体活动中，我注意发挥自己的特长，也学会了在团队中与他人配合协作。这些经历让我更加有责任感，也增强了我的组织协调能力和集体荣誉感。")] }),

        // 五、自我评价与今后努力方向
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("五、自我评价与今后努力方向")] }),
        new Paragraph({ children: [new TextRun({ text: "（一）自我评价。", bold: true, font }), new TextRun("总体而言，我认为自己是一个积极向上、努力进取的青年。我在思想上追求进步，坚决拥护党的领导，信仰坚定；在学习上勤奋刻苦，成绩较为优异；在生活上品行端正、乐于助人、热爱集体，具备一定的责任感和奉献精神。")] }),
        new Paragraph({ children: [new TextRun("同时，我也清醒地认识到自己还存在一些不足：一是理论学习还不够深入系统，有时停留在表面；二是自我要求还不够严格，遇到困难时偶尔会产生畏难情绪；三是主动服务意识和组织协调能力还有待进一步提高。")] }),
        new Paragraph({ children: [new TextRun({ text: "（二）今后努力方向。", bold: true, font }), new TextRun("针对以上不足，我为自己明确了努力方向：")] }),
        new Paragraph({ children: [new TextRun("第一，加强政治理论学习。持续深入学习党的创新理论和团的基本知识，用科学的理论武装头脑，不断提高政治觉悟和思想认识。")] }),
        new Paragraph({ children: [new TextRun("第二，严格要求自己。以团员标准检视日常言行，克服畏难情绪，在学习和生活中起到模范带头作用，做到自律、自强、自省。")] }),
        new Paragraph({ children: [new TextRun("第三，主动担当作为。积极参加团组织的各项活动，热心为同学和集体服务，在实践中历练本领、增长才干。")] }),
        new Paragraph({ children: [new TextRun("第四，全面提升素质。在学业上精益求精，在品德上严于律己，努力成长为德智体美劳全面发展的社会主义建设者和接班人。")] }),
        new Paragraph({ children: [new TextRun("如果组织能够吸纳我成为一名光荣的共青团员，我将倍加珍惜这份荣誉，不忘初心、砥砺前行，以实际行动践行入团誓言，不辜负组织的信任和培养；如果暂时未能如愿，我也绝不气馁，将继续查找差距、加倍努力，争取早日达到一名合格共青团员的标准。")] }),

        // 落款
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 400 }, children: [new TextRun({ text: "谈话人：　　　　　　", font })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "日期：　　年　　月　　日", font })] })
      ]
    }
  ]
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/workspace/tuan_talk/入团一对一谈话内容.docx", buffer);
  console.log("OK: 入团一对一谈话内容.docx");
});