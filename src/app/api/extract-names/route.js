import Anthropic from '@anthropic-ai/sdk';

export const config = { api: { bodyParser: { sizeLimit: '20mb' } } };

export async function POST(req) {
  try {
    const { images, texts } = await req.json();
    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

    const content = [];

    for (const img of (images || [])) {
      const [header, data] = img.split(',');
      const mediaType = header.match(/data:([^;]+)/)[1];
      content.push({ type: 'image', source: { type: 'base64', media_type: mediaType, data } });
    }

    if (texts && texts.length > 0) {
      content.push({ type: 'text', text: `Also extract names from this text:\n${texts.join('\n')}` });
    }

    content.push({
      type: 'text',
      text: `Extract ALL names of people from the above content.
Return ONLY a JSON array of name strings, nothing else. No markdown, no explanation.
Example: ["ישראל ישראלי", "שרה כהן", "מיכאל לוי"]
If no names found, return [].`
    });

    const response = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 2000,
      messages: [{ role: 'user', content }]
    });

    const text = response.content[0].text.trim();
    const names = JSON.parse(text);
    return Response.json({ names });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}
