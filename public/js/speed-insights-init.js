// Initialize Vercel Speed Insights
// This script imports and initializes Speed Insights for tracking web vitals
import { injectSpeedInsights } from './speed-insights.mjs';

// Initialize Speed Insights with default configuration
// debug: automatically enabled in development mode
// sampleRate: 1 means 100% of events are tracked
injectSpeedInsights({
  debug: false, // Set to true for debugging
  sampleRate: 1 // Track 100% of events
});
