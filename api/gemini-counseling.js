// api/gemini-counseling.js

export default async function handler(req, res) {
    // 1. Only allow POST requests
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed. Only POST requests are allowed.' });
    }

    const { title, url } = req.body;

    // 2. Validate input (at least URL is required)
    if (!url) {
        return res.status(400).json({ error: 'url parameter is required.' });
    }

    // 3. Read GEMINI_API_KEY from environment variables
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
        return res.status(500).json({ error: 'GEMINI_API_KEY environment variable is not configured.' });
    }

    // 4. Construct prompt securely without sharing sensitive personal info
    const prompt = `웹앱 이름: ${title || '이름 없음'}\n웹앱 주소: ${url}\n\n이 웹앱은 학교 수업 및 학급 운영 중 교사와 학생들이 활용하는 교육용 웹앱입니다. 이 웹앱의 이름과 주소 정보를 바탕으로, 해당 웹앱이 어떤 역할을 하고 무엇을 할 수 있는 앱인지 정중하고 명확한 한국어 한 문장(약 50자 내외, 최대 80자)의 '한 줄 소개'로 요약해 작성해 주세요. 요약 내용만 즉시 반환해야 하며, 따옴표나 서론("이 웹앱은 ~", "여기 소개입니다" 등)을 붙이지 마세요.`;

    // 5. Call Gemini API using built-in fetch
    // Using model gemini-3.0-flash-lite as requested
    const model = 'gemini-3.0-flash-lite';
    const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    try {
        const response = await fetch(geminiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                contents: [
                    {
                        parts: [
                            {
                                text: prompt
                            }
                        ]
                    }
                ],
                generationConfig: {
                    temperature: 0.2,
                    maxOutputTokens: 150
                }
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error('Gemini API Error:', errorText);
            return res.status(502).json({ error: 'Gemini API returned an error.', details: errorText });
        }

        const data = await response.json();
        
        // Parse content response
        const candidates = data.candidates;
        if (candidates && candidates.length > 0 && candidates[0].content && candidates[0].content.parts && candidates[0].content.parts.length > 0) {
            const textResult = candidates[0].content.parts[0].text.trim().replace(/^"|"$/g, ''); // strip optional quotes
            return res.status(200).json({ summary: textResult });
        } else {
            return res.status(502).json({ error: 'Invalid response format from Gemini API.', details: data });
        }
    } catch (err) {
        console.error('Serverless Function Exception:', err);
        return res.status(500).json({ error: 'Internal server error while calling Gemini API.', details: err.message });
    }
}
