import Anthropic from '@anthropic-ai/sdk';

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const SOURCES = [
  'bike-eu.com','cyclingindustry.news','nieuwsfiets.nu',
  'bicycleretailer.com','velobiz.de','bikebiz.com',
  'micromobilitybiz.com','zagdaily.com','tech.eu',
  'siliconcanals.com','road.cc'
];

const QUARTER_MONTHS = {
  Q1: { start:'January',   end:'March'     },
  Q2: { start:'April',     end:'June'      },
  Q3: { start:'July',      end:'September' },
  Q4: { start:'October',   end:'December'  }
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { quarter, year } = req.body;
  const prevYear = parseInt(year) - 1;
  const { start, end } = QUARTER_MONTHS[quarter];

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');

  const send = (type, data) => {
    res.write(`data: ${JSON.stringify({ type, ...data })}\n\n`);
  };

  try {
    send('status', { msg: `Scanning ${quarter} ${year} across ${SOURCES.length} sources…` });
    const currentDeals = await scan(quarter, year, start, end, false, send);

    send('status', { msg: `Found ${currentDeals.length} deals in ${quarter} ${year}. Now scanning ${quarter} ${prevYear} for comparison…` });
    const prevDeals = await scan(quarter, prevYear, start, end, true, send);

    send('status', { msg: `✓ Complete — ${currentDeals.length} deals in ${quarter} ${year} vs ${prevDeals.length} in ${quarter} ${prevYear}` });
    send('result', {
      currentDeals,
      prevDeals,
      currentCount: currentDeals.length,
      prevCount: prevDeals.length,
      quarter,
      year: parseInt(year),
      prevYear
    });
  } catch (err) {
    send('error', { msg: err.message });
  }

  res.end();
}

async function scan(quarter, year, start, end, lightMode, send) {
  const prompt = `Search for ALL M&A transactions and investment deals in the bike and micromobility industry from ${start} ${year} to ${end} ${year}. Include European and North American deals, with special attention to Dutch/Benelux companies.

Search these sources: ${SOURCES.join(', ')}

Include: acquisitions, mergers, VC rounds, PE investments, debt funding, recapitalisations, and bankruptcy sales.
${lightMode ? 'Return compact objects with just target, acquirer, segment, type fields.' : 'Return full details for each deal.'}

Return ONLY a raw JSON array, no markdown, no explanation. Each object:
{
  "target": "company name",
  "acquirer": "acquirer or investor name",
  "segment": one of "Bike Manufacturers"|"Service Providers"|"Retail / Distribution"|"PAC"|"Bike Leasing",
  "type": one of "Acquired"|"Invested in"|"Merged with"|"Recapitalised",
  "amount": "deal value or empty string",
  "description": "1 sentence about the deal",
  "source": "publication name",
  "dutch": true or false
}`;

  let fullText = '';

  try {
    const stream = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4000,
      tools: [{
        type: 'web_search_20250305',
        name: 'web_search',
        max_uses: 8,
        allowed_domains: SOURCES
      }],
      messages: [{ role: 'user', content: prompt }],
      stream: true
    });

    for await (const event of stream) {
      if (event.type === 'content_block_delta' && event.delta?.type === 'text_delta') {
        fullText += event.delta.text;
      }
      if (event.type === 'content_block_start' && event.content_block?.type === 'tool_use') {
        send('status', { msg: `Searching ${quarter} ${year}…` });
      }
    }
  } catch (e) {
    send('status', { msg: `Search error: ${e.message}` });
  }

  try {
    const s = fullText.indexOf('['), e = fullText.lastIndexOf(']');
    if (s !== -1 && e !== -1) return JSON.parse(fullText.slice(s, e + 1)).filter(d => d?.target);
  } catch (_) {}
  return [];
}
