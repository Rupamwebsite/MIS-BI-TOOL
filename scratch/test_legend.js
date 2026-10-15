const http = require('http');

function postJSON(path, payload) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload);
    const req = http.request({
      hostname: 'localhost',
      port: 5000,
      path: path,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data)
      }
    }, res => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => resolve(JSON.parse(body)));
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

async function testMultiSeries() {
  console.log('Testing Multi-Series Stacked Visual: Department by Bill Status...');
  const res = await postJSON('/api/datasets/4/visual', {
    xAxis: 'DEPARTMENT',
    yAxis: 'NET AMOUNT',
    groupBy: 'BILL STATUS',
    agg: 'sum',
    topN: 4
  });
  console.log('Labels (X-Axis):', res.visual?.labels);
  console.log('Legend Series (Datasets):', res.visual?.datasets?.map(d => ({ label: d.label, values: d.data })));
}

testMultiSeries().catch(console.error);
