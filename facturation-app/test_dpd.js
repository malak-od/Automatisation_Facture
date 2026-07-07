const path = require('path');
const registry = require('./src/registry');

async function run() {
  try {
    const carrier = registry.get('dpd');
    const files = { csv: [
      path.join(__dirname, 'uploads', '83f694738a760cfa6126dc011fbde3ba'),
      path.join(__dirname, 'uploads', '50bd483fb7393a5f4c0c1174789b145c'),
      path.join(__dirname, 'uploads', '69866cf2d3f129f25be05ede9bdfe6a4'),
      path.join(__dirname, 'uploads', '9ea0ffa31fddea8f45dac705bef920fd'),
      path.join(__dirname, 'uploads', '91f75e46303b4e2e431cff66ba094923'),
      path.join(__dirname, 'uploads', 'bcc25ec70acf42dd249c8ee575f7b3e6'),
      path.join(__dirname, 'uploads', 'a7e25afe4886364172d81c09165db553'),
      path.join(__dirname, 'uploads', 'f9b7eeb1d5f36c80e5087a94da385edb'),
      path.join(__dirname, 'uploads', '666cc34fc387dcabe8ab31a345d40d9c'),
      path.join(__dirname, 'uploads', 'c359585caa779047acd771175859c795'),
      path.join(__dirname, 'uploads', '3322c9f3f52bdb26515a00d8775b0522'),
    ] };
    const res = await carrier.process(files);
    console.log('Process result:', Object.keys(res));
  } catch (e) {
    console.error('Process threw:', e && e.stack ? e.stack : e);
    process.exit(1);
  }
}

run();
