const axios = require('axios');

const API_KEY = "EAApL-OLL!K0MTN_OlPHGCVUDePrJ_QMVsI2BEGOA~U$PWR;U.LCR0H+VZRES0LBD9X5A-Q;H1HcCYH?DqVAT%PNF:OZR%A;JvA%";

async function getToken() {
  const res = await axios.post(
    "https://trackapi.thailandpost.co.th/post/api/v1/authenticate/token",
    {},
    { headers: { Authorization: "Token " + API_KEY } }
  );
  return res.data.token;
}

async function trackParcel(trackingNumber) {
  const token = await getToken();
  const res = await axios.post(
    "https://trackapi.thailandpost.co.th/post/api/v1/track",
    {
      status: "all",
      language: "TH",
      barcode: [trackingNumber]
    },
    { headers: { Authorization: "Token " + token } }
  );
  return res.data;
}

module.exports = { trackParcel }; 