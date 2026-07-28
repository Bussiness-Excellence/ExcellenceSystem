export const config = {
  runtime: 'edge',
};

export default async function handler(req) {
  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), {
      status: 405,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  try {
    const { messages, contextData } = await req.json();

    if (!messages || !Array.isArray(messages)) {
      return new Response(JSON.stringify({ error: 'Missing or invalid messages array' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
    }



    // Build context string safely
    const contextStr = JSON.stringify(contextData).substring(0, 8000);

    const systemPrompt = `You are Excellence AI, an expert CRM data analyst.
You are helping a pharmaceutical manager analyze their dashboard data.
Answer briefly and professionally based ONLY on the following JSON data.
Current Data context: ${contextStr}`;

    // Format messages for Gemini API
    const geminiMessages = messages.map(msg => ({
      role: msg.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: msg.content }]
    }));

    // Inject system prompt into the first user message
    if (geminiMessages.length > 0 && geminiMessages[0].role === 'user') {
      geminiMessages[0].parts[0].text = `${systemPrompt}\n\nUser: ${geminiMessages[0].parts[0].text}`;
    }

    const payload = {
      contents: geminiMessages,
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 500,
      }
    };

    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${process.env.GEMINI_API_KEY}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await response.json();
    
    if (data.error) {
      throw new Error(data.error.message || 'Gemini API Error');
    }

    const replyText = data.candidates?.[0]?.content?.parts?.[0]?.text || 'Sorry, no valid response was returned.';

    return new Response(JSON.stringify({ reply: replyText }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error('AI Edge Function Error:', error);
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
    });
  }
}
