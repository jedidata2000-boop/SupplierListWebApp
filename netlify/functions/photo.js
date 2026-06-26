const { google } = require('googleapis');

exports.handler = async (event) => {
  const id = event.queryStringParameters && event.queryStringParameters.id;
  if (!id) return { statusCode: 400, body: 'Missing id' };

  try {
    const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
    const key = (process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || '').replace(/\\n/g, '\n');
    const auth = new google.auth.JWT(email, null, key, [
      'https://www.googleapis.com/auth/drive.readonly'
    ]);
    const drive = google.drive({ version: 'v3', auth });

    const res = await drive.files.get(
      { fileId: id, alt: 'media', supportsAllDrives: true },
      { responseType: 'stream' }
    );

    const buffer = await new Promise((resolve, reject) => {
      const chunks = [];
      res.data.on('data', chunk => chunks.push(chunk));
      res.data.on('end', () => resolve(Buffer.concat(chunks)));
      res.data.on('error', reject);
    });

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'image/jpeg',
        'Cache-Control': 'public, max-age=604800'
      },
      body: buffer.toString('base64'),
      isBase64Encoded: true
    };
  } catch (e) {
    return { statusCode: 404, body: 'Photo not found: ' + e.message };
  }
};
