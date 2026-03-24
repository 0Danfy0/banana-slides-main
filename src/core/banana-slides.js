/**
 * Banana Slides - AI PowerPoint Generator
 * AI 驱动的专业 PPT 生成器
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// 主题配置
const THEME = {
  PRIMARY: "005293",
  SECONDARY: "0078D4",
  ACCENT: "4472C4",
  DARK: "1E3A5F",
  TEXT: "2C3E50",
  WHITE: "FFFFFF",
  BG: "F5F8FC",
  GRAY: "7F8C8D",
};

/**
 * 加载配置文件
 */
function loadConfig() {
  const configPath = path.join(__dirname, "../../config.json");
  if (fs.existsSync(configPath)) {
    return JSON.parse(fs.readFileSync(configPath, "utf-8"));
  }
  return {};
}

/**
 * Banana Slides 主类
 */
class BananaSlides {
  constructor(options = {}) {
    this.config = loadConfig();
    this.options = options;

    // 初始化 pptxgenjs
    this.pres = new pptxgen();
    this.pres.layout = "LAYOUT_16x9";
    this.pres.title = options.title || "Banana Slides";
    this.pres.author = options.author || "Banana Slides";

    // 主题色
    this.colors = { ...THEME, ...(this.config.Theme || {}) };

    // 预设图标
    this.icons = {};

    // 输出目录
    this.outputDir = this.config.Output?.path || "./output";
    if (!fs.existsSync(this.outputDir)) {
      fs.mkdirSync(this.outputDir, { recursive: true });
    }
  }

  /**
   * 初始化图标
   */
  async initIcons(iconList = []) {
    const { iconToBase64Png } = require("../utils/icons");
    
    const defaultIcons = [
      "FaBullseye", "FaChartLine", "FaUsers", "FaBuilding",
      "FaCheckCircle", "FaShieldAlt", "FaCogs", "FaGlobe",
    ];

    for (const iconName of iconList.length > 0 ? iconList : defaultIcons) {
      try {
        this.icons[iconName] = await iconToBase64Png(iconName, "#" + this.colors.ACCENT, 256);
      } catch (e) {
        console.warn(`图标 ${iconName} 生成失败`);
      }
    }
  }

  /**
   * 添加封面页
   */
  addCoverSlide(title, subtitle, author, bgImage) {
    const s = this.pres.addSlide();
    
    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: "100%",
      fill: { color: this.colors.DARK },
    });

    s.addShape(this.pres.shapes.OVAL, {
      x: -1, y: -1, w: 4, h: 4,
      fill: { color: this.colors.SECONDARY, transparency: 70 },
    });
    s.addShape(this.pres.shapes.OVAL, {
      x: 8, y: 3, w: 3, h: 3,
      fill: { color: this.colors.ACCENT, transparency: 60 },
    });

    if (bgImage && fs.existsSync(bgImage)) {
      s.addImage({ path: bgImage, x: 0, y: 0, w: "100%", h: "100%", transparency: 80 });
    }

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: 0.15,
      fill: { color: this.colors.ACCENT },
    });

    s.addText(subtitle || "", {
      x: 0.5, y: 1.3, w: 9, h: 0.7,
      fontSize: 22, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, align: "center", charSpacing: 6,
    });

    s.addText(title, {
      x: 0.5, y: 2, w: 9, h: 1,
      fontSize: 44, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, align: "center",
    });

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 3.5, y: 4, w: 3, h: 0.03,
      fill: { color: this.colors.ACCENT },
    });

    s.addText(author || "", {
      x: 0.5, y: 4.3, w: 9, h: 0.5,
      fontSize: 16, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, align: "center",
    });

    return s;
  }

  /**
   * 添加目录页
   */
  addTocSlide(items) {
    const s = this.pres.addSlide();
    s.background = { color: this.colors.BG };

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: 1.1,
      fill: { color: this.colors.PRIMARY },
    });
    s.addText("目 录", {
      x: 0.5, y: 0.3, w: 9, h: 0.5,
      fontSize: 28, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, margin: 0,
    });

    items.forEach((item, i) => {
      const x = 0.5 + (i % 3) * 3.1;
      const y = 1.4 + Math.floor(i / 3) * 1.9;

      s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 2.9, h: 1.7,
        fill: { color: this.colors.WHITE },
        shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "000000", opacity: 0.1 },
        rectRadius: 0.1,
      });

      s.addShape(this.pres.shapes.RECTANGLE, {
        x, y: y + 0.3, w: 0.08, h: 1.1,
        fill: { color: this.colors.ACCENT },
      });

      s.addText(item.num || String(i + 1).padStart(2, "0"), {
        x: x + 0.2, y: y + 0.15, w: 0.8, h: 0.6,
        fontSize: 30, fontFace: "Arial",
        color: this.colors.ACCENT, bold: true, margin: 0,
      });

      s.addText(item.title, {
        x: x + 0.2, y: y + 0.75, w: 2.5, h: 0.5,
        fontSize: 12, fontFace: "Microsoft YaHei",
        color: this.colors.TEXT, bold: true, margin: 0,
      });

      if (item.desc) {
        s.addText(item.desc, {
          x: x + 0.2, y: y + 1.2, w: 2.5, h: 0.3,
          fontSize: 10, fontFace: "Microsoft YaHei",
          color: this.colors.GRAY, margin: 0,
        });
      }
    });

    return s;
  }

  /**
   * 添加章节分隔页
   */
  addSectionSlide(num, title, eng, bgColor) {
    const s = this.pres.addSlide();
    s.background = { color: bgColor || this.colors.PRIMARY };

    s.addText(num, {
      x: -0.5, y: 0.3, w: 5, h: 4,
      fontSize: 220, fontFace: "Arial",
      color: this.colors.SECONDARY, bold: true, margin: 0,
    });

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0.5, y: 2.4, w: 2, h: 0.05,
      fill: { color: this.colors.WHITE },
    });

    s.addText(title, {
      x: 0.5, y: 2.6, w: 8, h: 1,
      fontSize: 38, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, margin: 0,
    });

    s.addText(eng || "", {
      x: 0.5, y: 3.5, w: 8, h: 0.5,
      fontSize: 12, fontFace: "Arial",
      color: this.colors.WHITE, charSpacing: 3, margin: 0,
    });

    s.addShape(this.pres.shapes.OVAL, {
      x: 8, y: 3.3, w: 2.5, h: 2.5,
      fill: { color: this.colors.SECONDARY, transparency: 50 },
    });

    return s;
  }

  /**
   * 添加图标卡片布局
   */
  addIconCardsSlide(title, cards) {
    const s = this.pres.addSlide();
    s.background = { color: this.colors.BG };

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: 0.9,
      fill: { color: this.colors.PRIMARY },
    });
    s.addText(title, {
      x: 0.5, y: 0.2, w: 9, h: 0.5,
      fontSize: 24, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, margin: 0,
    });

    cards.forEach((card, i) => {
      const x = 0.5 + i * 2.35;
      const h = card.height || 2.9;

      s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.1, w: 2.2, h,
        fill: { color: this.colors.WHITE },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.1 },
        rectRadius: 0.1,
      });

      s.addShape(this.pres.shapes.OVAL, {
        x: x + 0.6, y: 1.3, w: 1, h: 1,
        fill: { color: this.colors.ACCENT, transparency: 15 },
      });

      if (card.icon && this.icons[card.icon]) {
        s.addImage({ data: this.icons[card.icon], x: x + 0.75, y: 1.45, w: 0.7, h: 0.7 });
      }

      s.addText(card.title, {
        x: x + 0.1, y: 2.4, w: 2, h: 0.5,
        fontSize: 14, fontFace: "Microsoft YaHei",
        color: this.colors.PRIMARY, bold: true, align: "center", margin: 0,
      });

      s.addText(card.desc, {
        x: x + 0.1, y: 2.9, w: 2, h: card.descHeight || 1,
        fontSize: 10, fontFace: "Microsoft YaHei",
        color: this.colors.TEXT, align: "center", margin: 0,
      });
    });

    return s;
  }

  /**
   * 添加双栏布局
   */
  addTwoColumnSlide(title, left, right) {
    const s = this.pres.addSlide();
    s.background = { color: this.colors.BG };

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: 0.9,
      fill: { color: this.colors.PRIMARY },
    });
    s.addText(title, {
      x: 0.5, y: 0.2, w: 9, h: 0.5,
      fontSize: 24, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, margin: 0,
    });

    s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: 1.1, w: 4.3, h: 4.3,
      fill: { color: left.bgColor || this.colors.PRIMARY },
      rectRadius: 0.1,
    });
    s.addText(left.title, {
      x: 0.7, y: 1.3, w: 3.9, h: 0.5,
      fontSize: 18, fontFace: "Microsoft YaHei",
      color: left.textColor || this.colors.WHITE, bold: true, margin: 0,
    });
    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0.7, y: 1.85, w: 1.5, h: 0.03,
      fill: { color: left.textColor || this.colors.WHITE, transparency: 50 },
    });
    s.addText(left.content, {
      x: 0.7, y: 2.0, w: 3.9, h: 3.2,
      fontSize: 12, fontFace: "Microsoft YaHei",
      color: left.textColor || this.colors.WHITE, margin: 0,
    });

    s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.1, w: 4.3, h: 4.3,
      fill: { color: right.bgColor || this.colors.WHITE },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.1 },
      rectRadius: 0.1,
    });
    s.addText(right.title, {
      x: 5.4, y: 1.3, w: 3.9, h: 0.5,
      fontSize: 18, fontFace: "Microsoft YaHei",
      color: this.colors.PRIMARY, bold: true, margin: 0,
    });
    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 5.4, y: 1.85, w: 1.5, h: 0.03,
      fill: { color: this.colors.ACCENT },
    });
    s.addText(right.content, {
      x: 5.4, y: 2.0, w: 3.9, h: 3.2,
      fontSize: 12, fontFace: "Microsoft YaHei",
      color: this.colors.TEXT, margin: 0,
    });

    return s;
  }

  /**
   * 添加组织架构图
   */
  addOrgChartSlide(title, root, children) {
    const s = this.pres.addSlide();
    s.background = { color: this.colors.BG };

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: 0.9,
      fill: { color: this.colors.PRIMARY },
    });
    s.addText(title, {
      x: 0.5, y: 0.2, w: 9, h: 0.5,
      fontSize: 24, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, margin: 0,
    });

    s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, {
      x: 3.5, y: 1.1, w: 3, h: 0.8,
      fill: { color: this.colors.PRIMARY },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.15 },
      rectRadius: 0.08,
    });
    s.addText(root, {
      x: 3.5, y: 1.25, w: 3, h: 0.5,
      fontSize: 12, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, align: "center", margin: 0,
    });

    const lineX = 5;
    s.addShape(this.pres.shapes.LINE, { x: lineX, y: 1.9, w: 0, h: 0.3, line: { color: this.colors.ACCENT, width: 2 } });
    s.addShape(this.pres.shapes.LINE, { x: 2.5, y: 2.2, w: 5, h: 0, line: { color: this.colors.ACCENT, width: 2 } });

    children.forEach((child, i) => {
      const childW = children.length > 3 ? 2.5 : 2.8;
      const startX = children.length === 3 ? 0.7 : 1.5;
      const x = startX + i * (childW + 0.3);

      s.addShape(this.pres.shapes.LINE, { x: x + childW / 2, y: 2.2, w: 0, h: 0.2, line: { color: this.colors.ACCENT, width: 2 } });

      s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.4, w: childW, h: 1.5,
        fill: { color: this.colors.SECONDARY },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.12 },
        rectRadius: 0.08,
      });
      s.addText(child.title, {
        x, y: 2.5, w: childW, h: 0.5,
        fontSize: 11, fontFace: "Microsoft YaHei",
        color: this.colors.WHITE, bold: true, align: "center", margin: 0,
      });
      s.addText(child.desc, {
        x, y: 3.0, w: childW, h: 0.8,
        fontSize: 10, fontFace: "Microsoft YaHei",
        color: this.colors.WHITE, align: "center", margin: 0,
      });
    });

    return s;
  }

  /**
   * 添加列表详情页
   */
  addListDetailSlide(title, items) {
    const s = this.pres.addSlide();
    s.background = { color: this.colors.BG };

    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: 0.9,
      fill: { color: this.colors.PRIMARY },
    });
    s.addText(title, {
      x: 0.5, y: 0.2, w: 9, h: 0.5,
      fontSize: 24, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, margin: 0,
    });

    const itemH = 4.3 / items.length - 0.1;
    items.forEach((item, i) => {
      const y = 1.1 + i * (itemH + 0.1);

      s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.5, y, w: 9, h: itemH,
        fill: { color: this.colors.WHITE },
        shadow: { type: "outer", blur: 4, offset: 1, angle: 135, color: "000000", opacity: 0.08 },
        rectRadius: 0.08,
      });

      s.addShape(this.pres.shapes.RECTANGLE, {
        x: 0.5, y: y + 0.2, w: 0.08, h: itemH - 0.4,
        fill: { color: this.colors.ACCENT },
      });

      s.addText(item.title, {
        x: 0.8, y: y + 0.1, w: 8.5, h: 0.4,
        fontSize: item.titleSize || 14, fontFace: "Microsoft YaHei",
        color: this.colors.PRIMARY, bold: true, margin: 0,
      });

      s.addText(item.desc, {
        x: 0.8, y: y + 0.5, w: 8.5, h: itemH - 0.6,
        fontSize: item.descSize || 11, fontFace: "Microsoft YaHei",
        color: this.colors.TEXT, margin: 0,
      });
    });

    return s;
  }

  /**
   * 添加结束页
   */
  addEndSlide(text, subtext) {
    const s = this.pres.addSlide();
    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: "100%",
      fill: { color: this.colors.DARK },
    });
    s.addShape(this.pres.shapes.OVAL, {
      x: -1, y: -1, w: 4, h: 4,
      fill: { color: this.colors.SECONDARY, transparency: 70 },
    });
    s.addShape(this.pres.shapes.OVAL, {
      x: 7, y: 3, w: 4, h: 4,
      fill: { color: this.colors.ACCENT, transparency: 60 },
    });
    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: "100%", h: 0.15,
      fill: { color: this.colors.ACCENT },
    });
    s.addText(text || "感谢聆听", {
      x: 0.5, y: 2, w: 9, h: 1.2,
      fontSize: 56, fontFace: "Microsoft YaHei",
      color: this.colors.WHITE, bold: true, align: "center",
    });
    s.addShape(this.pres.shapes.RECTANGLE, {
      x: 3.5, y: 3.4, w: 3, h: 0.03,
      fill: { color: this.colors.ACCENT },
    });
    if (subtext) {
      s.addText(subtext, {
        x: 0.5, y: 3.7, w: 9, h: 0.6,
        fontSize: 24, fontFace: "Microsoft YaHei",
        color: this.colors.WHITE, align: "center",
      });
    }

    return s;
  }

  /**
   * 保存 PPT
   */
  async save(filename) {
    const outputPath = filename || path.join(this.outputDir, this.config.Output?.defaultFilename || "presentation.pptx");
    await this.pres.writeFile({ fileName: outputPath });
    console.log(`PPT 已保存: ${outputPath}`);
    console.log(`共 ${this.pres.slides.length} 页`);
    return outputPath;
  }

  get slideCount() {
    return this.pres.slides.length;
  }
}

module.exports = BananaSlides;
