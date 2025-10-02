const { app } = require('@azure/functions');
require('./functions/BirthdayImage.js');

app.setup({
    enableHttpStream: true,
});
