import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 5000;

// JSON body parser with 50MB limit for large report payloads
app.use(express.json({ limit: '50mb' }));

// CORS headers to allow Vite client on port 3000 to call the server
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, PATCH, DELETE');
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,Content-Type,X-Filename');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// Ensure storage folders exist
const DB_DIR = path.join(__dirname, 'local_db');
const UPLOADS_DIR = path.join(__dirname, 'uploads');

if (!fs.existsSync(DB_DIR)) {
  fs.mkdirSync(DB_DIR, { recursive: true });
}
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

// Serve uploaded images/PDFs statically
app.use('/uploads', express.static(UPLOADS_DIR));

// 1. Save Report Data
app.post('/api/save-data', (req, res) => {
  const { id, payload } = req.body;
  if (!id) {
    return res.status(400).json({ error: 'Missing ID' });
  }
  try {
    const filePath = path.join(DB_DIR, `${id}.json`);
    fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), 'utf-8');
    console.log(`[DB] Saved data for ID: ${id}`);
    res.json({ success: true });
  } catch (error) {
    console.error('Failed to save local data:', error);
    res.status(500).json({ error: 'Failed to save data' });
  }
});

// 2. Load Report Data
app.get('/api/load-data/:id', (req, res) => {
  const { id } = req.params;
  try {
    const filePath = path.join(DB_DIR, `${id}.json`);
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf-8');
      res.json(JSON.parse(raw));
    } else {
      res.status(404).json({ error: 'Not found' });
    }
  } catch (error) {
    console.error('Failed to load local data:', error);
    res.status(500).json({ error: 'Failed to load data' });
  }
});

// 3. Raw Upload for Images/PDFs (Zero-dependency binary handler)
app.post(
  '/api/upload-image',
  express.raw({ type: ['image/jpeg', 'image/png', 'image/gif', 'application/pdf'], limit: '20mb' }),
  (req, res) => {
    const contentType = req.headers['content-type'];
    const clientFilename = req.headers['x-filename'];
    
    let extension = 'jpg';
    if (contentType === 'application/pdf') extension = 'pdf';
    else if (contentType === 'image/png') extension = 'png';
    else if (contentType === 'image/gif') extension = 'gif';

    const filename = clientFilename 
      ? String(clientFilename) 
      : `upload_${Date.now()}.${extension}`;

    const filePath = path.join(UPLOADS_DIR, filename);

    try {
      fs.writeFileSync(filePath, req.body);
      
      // Serve via actual server IP or hostname dynamically
      const host = req.headers.host || `localhost:${PORT}`;
      const url = `http://${host.split(':')[0]}:${PORT}/uploads/${filename}`;
      
      console.log(`[Upload] File saved to ${filePath} -> ${url}`);
      res.json({ url });
    } catch (error) {
      console.error('Upload failed:', error);
      res.status(500).json({ error: 'Failed to write upload file' });
    }
  }
);

// 4. Dynamic Ollama AI Proxy (handles CORS & automatic model selection)
app.post('/api/ollama-chat', async (req, res) => {
  const { prompt } = req.body;
  if (!prompt) {
    return res.status(400).json({ error: 'Missing prompt' });
  }

  // A. List local Ollama models and find the best match
  let modelToUse = 'gemma2:9b'; // Default fallback
  try {
    const tagsRes = await fetch('http://localhost:11434/api/tags');
    if (tagsRes.ok) {
      const tagsData = await tagsRes.json();
      const models = tagsData.models || [];
      if (models.length > 0) {
        const modelNames = models.map(m => m.name);
        // Preference list: gemma2, qwen2.5, llama3.1, llama3, etc.
        const preferred = ['gemma2:9b', 'gemma2', 'qwen2.5:7b-instruct', 'qwen2.5', 'llama3.1', 'llama3', 'llama2', 'mistral'];
        const found = preferred.find(pref => modelNames.some(name => name.startsWith(pref)));
        if (found) {
          const exactMatch = modelNames.find(name => name.startsWith(found));
          modelToUse = exactMatch;
        } else {
          // If no preferred model, use the first model found
          modelToUse = modelNames[0];
        }
      }
    }
  } catch (err) {
    console.warn('[Ollama] Could not query models list, using default model: gemma2:9b');
  }

  console.log(`[Ollama] Selected model for request: "${modelToUse}"`);

  // B. Call local Ollama chat completions endpoint
  try {
    const ollamaResponse = await fetch('http://localhost:11434/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: modelToUse,
        messages: [{ role: 'user', content: prompt }],
        response_format: { type: 'json_object' },
        temperature: 0.1 // low temperature for strict json output consistency
      })
    });

    if (!ollamaResponse.ok) {
      const errText = await ollamaResponse.text();
      return res.status(ollamaResponse.status).json({ error: `Ollama error: ${errText}` });
    }

    const result = await ollamaResponse.json();
    const textResult = result.choices?.[0]?.message?.content || '{}';
    res.json({ text: textResult, model: modelToUse });
  } catch (error) {
    console.error('[Ollama] Proxy request failed:', error);
    res.status(500).json({ error: 'Could not connect to Ollama. Make sure Ollama is running on your PC (http://localhost:11434).' });
  }
});

// 5. Statically serve the React build (for high-performance single-port local production)
const DIST_DIR = path.join(__dirname, 'dist');
if (fs.existsSync(DIST_DIR)) {
  app.use(express.static(DIST_DIR));
  app.get('*', (req, res, next) => {
    // API and uploads routes should bypass index.html static serving
    if (req.path.startsWith('/api/') || req.path.startsWith('/uploads/')) {
      return next();
    }
    res.sendFile(path.join(DIST_DIR, 'index.html'));
  });
  console.log(`[Static] Serving React production build from ./dist`);
}

app.listen(PORT, '0.0.0.0', () => {
  console.log(`=================================================`);
  console.log(`🏠 Weekly Report Local Server is running on:`);
  console.log(`   - Local PC:   http://localhost:${PORT}`);
  console.log(`   - LAN URL:    http://0.0.0.0:${PORT} (Access from mobile!)`);
  console.log(`=================================================`);
});
