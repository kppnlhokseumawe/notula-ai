import { json } from '@sveltejs/kit';
import {GOOGLE_API_KEY} from '$env/static/private';

const API_URL =
  'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';

export async function POST({ request }) {
  try {
    const { prompt } = await request.json();
    if (!prompt) {
      return json({ error: 'Prompt kosong' }, { status: 400 });
    }

    const response = await fetch(
      `${API_URL}?key=${GOOGLE_API_KEY}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }]
        })
      }
    );

    console.log('API :', `${API_URL}?key=${GOOGLE_API_KEY}`);
    
    const result = await response.json();
    console.log('Request Body :', result)
    const text =
      result?.candidates?.[0]?.content?.parts?.[0]?.text || '';

    return json({ text });
  } catch (error) {
    console.error(error);
    return json({ error: 'Gagal generate' }, { status: 500 });
  }
}
