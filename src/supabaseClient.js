import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://tsxhjrbesjyhwgtsqzzo.supabase.co'
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRzeGhqcmJlc2p5aHdndHNxenpvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzcyNDk5ODQsImV4cCI6MjA5MjgyNTk4NH0.GyKq_oRsjZ7SgatmCKf06mQ7-OQwlh1LFX9IHgs2vnM'

export const supabase = createClient(supabaseUrl, supabaseKey)
