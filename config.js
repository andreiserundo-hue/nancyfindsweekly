// Loads secrets from env vars or a local (gitignored) secrets.json.
// Copy secrets.example.json -> secrets.json and fill in your values.
let s={}; try{ s=require('./secrets.json'); }catch(e){}
module.exports={
  SHOP_TOKEN: process.env.SHOP_TOKEN || s.SHOP_TOKEN || '',
  SHOP:       process.env.SHOP       || s.SHOP       || '',
  GLUED_KEY:  process.env.GLUED_KEY  || s.GLUED_KEY  || '',
  WS:         process.env.WS         || s.WS         || '',
};
