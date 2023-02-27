const size = 7;
const border = {
  top: {style:'thin'},
  left: {style:'thin'},
  bottom: {style:'thin'},
  right: {style:'thin'},
}
const fontName = 'DejaVu Sans';

const boldCenter = {
  font: {
    bold: true,
    name: fontName,
    size,
  },
  alignment: {
    wrapText: true,
    horizontal: 'center', 
    vertical: 'middle',
  },
  border,
}

const verticalBold = {
  font: {
    bold: true,
    name: fontName,
    size,
  },
  alignment: {
    horizontal: 'center', 
    vertical: 'middle',
    textRotation: 90,
    wrapText: true,
  },
  border,
}

const right = {
  font: {
    name: fontName,
    size,
  },
  alignment: {
    horizontal: 'right', 
    vertical: 'middle',
  },
  border,
}

const rightBold = {
  font: {
    bold: true,
    name: fontName,
    size,
  },
  alignment: {
    horizontal: 'right', 
    vertical: 'middle',
  },
  border,
}

const bold = {
  font: {
    bold: true,
    name: fontName,
    size,
  },
  alignment: {
    vertical: 'middle',
  },
  border,
}

const center = {
  font: {
    name: fontName,
    size,
  },
  alignment: {
    horizontal: 'center', 
    vertical: 'middle',
  },
  border,
}

module.exports = {
  boldCenter,
  verticalBold,
  right,
  bold,
  center,
  rightBold,
}