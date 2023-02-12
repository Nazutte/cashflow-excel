const boldCenter = {
  font: {
    bold: true,
    name: 'Calibri',
    size: 10,
  },
  alignment: {
    horizontal: 'center', 
    vertical: 'middle',
  }
}

const boldCenterFill = {
  font: {
    bold: true,
    name: 'Calibri',
    size: 10,
  },
  alignment: {
    horizontal: 'center', 
    vertical: 'middle',
  },
  fill: {
    type: 'pattern',
    pattern:'solid',
    bgColor:{argb: '6398c1'},
  }
}

module.exports = {
  boldCenter,
  boldCenterFill,
}