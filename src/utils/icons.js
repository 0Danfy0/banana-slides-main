/**
 * Banana Slides - Icon Renderer
 * 图标渲染工具
 */

const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Font Awesome
const {
  FaCogs, FaProjectDiagram, FaUsers, FaChartLine, FaCheckCircle,
  FaBuilding, FaLightbulb, FaBullseye, FaHandshake, FaShieldAlt,
  FaClock, FaArrowRight, FaCircle, FaGlobe, FaStar, FaRocket,
  FaBriefcase, FaDatabase, FaCog, FaFileAlt, FaChartBar, FaClipboardList
} = require("react-icons/fa");

// Material Design
const {
  MdAssessment, MdBuild, MdPeople, MdBusiness, MdTrendingUp,
  MdSecurity, MdSpeed, MdStar, MdSettings, MdAnalytics
} = require("react-icons/md");

// Box Icon
const {
  BiBriefcase, BiTargetLock, BiNetworkChart, BiPlus, BiMinus,
  BiCaretRight, BiChart, BiCode, BiData, BiPlanet
} = require("react-icons/bi");

const ICONS = {
  FaCogs, FaProjectDiagram, FaUsers, FaChartLine, FaCheckCircle,
  FaBuilding, FaLightbulb, FaBullseye, FaHandshake, FaShieldAlt,
  FaClock, FaArrowRight, FaCircle, FaGlobe, FaStar, FaRocket,
  FaBriefcase, FaDatabase, FaCog, FaFileAlt, FaChartBar, FaClipboardList,
  MdAssessment, MdBuild, MdPeople, MdBusiness, MdTrendingUp,
  MdSecurity, MdSpeed, MdStar, MdSettings, MdAnalytics,
  BiBriefcase, BiTargetLock, BiNetworkChart, BiPlus, BiMinus,
  BiCaretRight, BiChart, BiCode, BiData, BiPlanet,
};

function renderIconSvg(iconName, color = "#000000", size = 256) {
  const Icon = ICONS[iconName];
  if (!Icon) return null;
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color, size: String(size) })
  );
}

async function iconToBase64Png(iconName, color, size = 256) {
  try {
    const svg = renderIconSvg(iconName, color, size);
    if (!svg) return null;
    const buffer = await sharp(Buffer.from(svg)).png().toBuffer();
    return "image/png;base64," + buffer.toString("base64");
  } catch (error) {
    return null;
  }
}

async function batchGenerateIcons(iconNames, color, size = 256) {
  const icons = {};
  for (const name of iconNames) {
    icons[name] = await iconToBase64Png(name, color, size);
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  return icons;
}

function getAvailableIcons() {
  return Object.keys(ICONS);
}

module.exports = { ICONS, renderIconSvg, iconToBase64Png, batchGenerateIcons, getAvailableIcons };
