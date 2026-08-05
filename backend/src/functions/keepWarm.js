const { app } = require('@azure/functions');

// Fires every 5 minutes, forever — does nothing but log a line. Its only job is
// to stop the Consumption plan from ever fully spinning down between real
// requests, so users don't hit a multi-second cold start after sign-in.
app.timer('keepWarm', {
  schedule: '0 */5 * * * *',
  handler: async (myTimer, context) => {
    context.log('Keep-warm ping');
  }
});
