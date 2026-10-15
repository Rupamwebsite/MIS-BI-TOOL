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
      res.on('end', () => {
        try {
          resolve(JSON.parse(body));
        } catch (e) {
          resolve(body);
        }
      });
    });

    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

async function run() {
  console.log('--- TEST 1: Ask BI Copilot on Dataset 4 (ipdPatientReport 4,663 rows) ---');
  const askRes = await postJSON('/api/datasets/4/ask', { query: 'Top 5 doctors by net amount' });
  console.log('Success:', askRes.success);
  console.log('X-Axis (Dimension):', askRes.xAxis);
  console.log('Y-Axis (Metric):', askRes.yAxis);
  console.log('Aggregation:', askRes.agg);
  console.log('Answer:', askRes.answer);
  console.log('Top items:', askRes.visual?.items?.slice(0, 5));

  console.log('\n--- TEST 2: Visual Studio Custom Visual (DEPARTMENT vs NET AMOUNT) ---');
  const visRes = await postJSON('/api/datasets/4/visual', {
    xAxis: 'DEPARTMENT',
    yAxis: 'NET AMOUNT',
    agg: 'sum',
    topN: 5,
    sort: 'desc'
  });
  console.log('Success:', visRes.success);
  console.log('Visual Labels:', visRes.visual?.labels);
  console.log('Visual Dataset 0 Data:', visRes.visual?.datasets?.[0]?.data);
  console.log('Items Count:', visRes.visual?.items?.length);
}

run().catch(console.error);
