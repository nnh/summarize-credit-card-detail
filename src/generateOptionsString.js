function generateOptionsString_(inputParameters = new Map()) {
  const portrait = inputParameters.has('portrait')
    ? `&portrait=${inputParameters.get('portrait')}`
    : '&portrait=true';
  const scale = inputParameters.has('scale')
    ? `&portrait=${inputParameters.get('scale')}`
    : '&scale=2';
  const printtitle = inputParameters.has('printtitle')
    ? `&printtitle=${inputParameters.get('printtitle')}`
    : '&printtitle=false';
  const sheetnames = inputParameters.has('sheetnames')
    ? `&printtitle=${inputParameters.get('sheetnames')}`
    : '&sheetnames=false';
  const gridlines = inputParameters.has('gridlines')
    ? `&printtitle=${inputParameters.get('gridlines')}`
    : '&gridlines=false';
  const gid = inputParameters.has('gid')
    ? `&gid=${inputParameters.get('gid')}`
    : '';
  const options =
    '&exportFormat=pdf&format=pdf' +
    '&size=A4' + // Paper size (A4)
    gid +
    portrait + // true: Portrait, false: Landscape
    scale + // 1:100%, 2:fit to width, 3:fit to height, 4:fit to page
    '&top_margin=0.50' + // Top margin
    '&right_margin=0.50' + // Right margin
    '&bottom_margin=0.50' + // Bottom margin
    '&left_margin=0.50' + // Left margin
    '&horizontal_alignment=CENTER' + // Horizontal alignment
    '&vertical_alignment=TOP' + // Vertical alignment
    printtitle + // Show sheet name
    sheetnames + // Show sheet names
    gridlines + // Show gridlines
    '&fzr=true' + // Show frozen rows
    '&fzc=true'; // Show frozen columns
  return options;
}