const { Client } = require('pg');

const regions = [
  'ap-northeast-2', 'ap-northeast-1', 'ap-northeast-3',
  'ap-southeast-1', 'ap-southeast-2', 'ap-south-1',
  'us-east-1', 'us-east-2', 'us-west-1', 'us-west-2',
  'eu-west-1', 'eu-west-2', 'eu-west-3', 'eu-central-1',
  'sa-east-1', 'ca-central-1'
];

async function run() {
  let client = null;
  for (const region of regions) {
    const host = `aws-0-${region}.pooler.supabase.com`;
    const c = new Client({
      connectionString: `postgresql://postgres.xootqaeuixpsszcejhev:chongmuinsa2027@${host}:6543/postgres`,
      connectionTimeoutMillis: 5000
    });
    try {
      await c.connect();
      console.log('Connected successfully on', region);
      client = c;
      break;
    } catch(e) {
      if (e.message.includes('tenant/user')) {
        // Pooler exists, but project not in this region
        console.log(`Region ${region}: Not found`);
      } else {
        console.log(`Region ${region}: Error`, e.message);
      }
    }
  }

  if (!client) {
    console.log("Failed to find correct region.");
    process.exit(1);
  }

  try {
    await client.query(`
      CREATE TABLE IF NOT EXISTS public.reports (
        id text PRIMARY KEY,
        data jsonb,
        "lastSaved" text,
        status text,
        updated_at timestamp with time zone DEFAULT timezone('utc'::text, now()) NOT NULL
      );
    `);
    console.log("Created reports table");

    await client.query(`ALTER TABLE public.reports DISABLE ROW LEVEL SECURITY;`);
    console.log("Disabled RLS on reports");

    await client.query(`
      INSERT INTO storage.buckets (id, name, public)
      VALUES ('images', 'images', true)
      ON CONFLICT (id) DO UPDATE SET public = true;
    `);
    console.log("Created images bucket");

    await client.query(`CREATE POLICY "Public Access" ON storage.objects FOR SELECT USING (bucket_id = 'images');`).catch(e => {});
    await client.query(`CREATE POLICY "Public Insert" ON storage.objects FOR INSERT WITH CHECK (bucket_id = 'images');`).catch(e => {});
    console.log("Storage policies created");

  } catch(e) {
    console.error("Setup Error:", e);
  } finally {
    await client.end();
  }
}

run();
