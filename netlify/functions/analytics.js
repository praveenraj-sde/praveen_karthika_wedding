// Proxy for visitor analytics — forwards to the same GAS Web App as RSVP.
// Returns 204 immediately to the client regardless of GAS outcome.

const GAS_URL =
  'https://script.google.com/macros/s/AKfycbz7rE1PaMqUNLQY1Sv5VlRs4u2PwAn8WIcBtYSeUldoN0xu9CoTOSGDgpPLohMK2tBwbw/exec';

export const handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 204, body: '' };
  }

  try {
    await fetch(GAS_URL, {
      method:   'POST',
      headers:  { 'Content-Type': 'text/plain;charset=utf-8' },
      body:     event.body,
      redirect: 'follow',
      signal:   AbortSignal.timeout(8000),
    });
  } catch (err) {
    // Silently swallow — analytics must never affect the visitor's experience
    console.error('Analytics GAS error:', err.message);
  }

  return { statusCode: 204, body: '' };
};
