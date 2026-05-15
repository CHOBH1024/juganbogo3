import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://xootqaeuixpsszcejhev.supabase.co';
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inhvb3RxYWV1aXhwc3N6Y2VqaGV2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg3ODk3MDgsImV4cCI6MjA5NDM2NTcwOH0.W2h7M1zUZFNG6KjtQm92CfG3ixcllhhW2_Az6loxYJI';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

