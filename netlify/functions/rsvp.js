// Proxy between the browser and Google Apps Script.
// GAS drops the CORS header on its redirect — calling it server-side avoids the issue.

const GAS_URL =
  'https://script.google.com/macros/s/AKfycbz7rE1PaMqUNLQY1Sv5VlRs4u2PwAn8WIcBtYSeUldoN0xu9CoTOSGDgpPLohMK2tBwbw/exec';

const CORS = {
  'Access-Control-Allow-Origin':  'https://praveen-karthika-wedding.netlify.app',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export const handler = async (event) => {
  // Preflight
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: CORS, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify({ success: false, error: 'Method not allowed' }),
    };
  }

  try {
    // Forward to GAS using text/plain to avoid a second preflight there
    const gasRes = await fetch(GAS_URL, {
      method:   'POST',
      headers:  { 'Content-Type': 'text/plain;charset=utf-8' },
      body:     event.body,
      redirect: 'follow',
    });

    const text = await gasRes.text();

    return {
      statusCode: 200,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: text,
    };
  } catch (err) {
    console.error('RSVP proxy error:', err);
    return {
      statusCode: 500,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify({ success: false, error: 'Server error — please try again' }),
    };
  }
};
