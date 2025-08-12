import OSS from 'ali-oss';
import ExcelJS from 'exceljs';
import { v4 as uuidv4 } from 'uuid';

const client = new OSS({
  region: process.env.OSS_REGION,       // 例：oss-cn-hangzhou
  accessKeyId: process.env.OSS_ACCESS_KEY_ID,
  accessKeySecret: process.env.OSS_ACCESS_KEY_SECRET,
  bucket: process.env.OSS_BUCKET
});

export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed, use POST.' });
    }

    const { file_name, sheet_name = 'Sheet1', headers, rows } = req.body || {};

    if (!file_name || !headers || !rows) {
      return res.status(400).json({ error: 'Missing required fields: file_name, headers, rows' });
    }

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet(sheet_name);
    ws.addRow(headers);
    rows.forEach(row => ws.addRow(row));

    const buffer = await workbook.xlsx.writeBuffer();

    const ossKey = `coze-excels/${uuidv4()}_${file_name.replace(/[^a-zA-Z0-9_.-]/g, '_')}`;

    await client.put(ossKey, Buffer.from(buffer));

    const downloadUrl = `https://${process.env.OSS_BUCKET}.${process.env.OSS_REGION}.aliyuncs.com/${ossKey}`;

    res.status(200).json({
      status: 'success',
      download_url: downloadUrl
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || String(err) });
  }
}
