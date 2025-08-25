import { google } from "googleapis";

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json", // service account key
  scopes: ["https://www.googleapis.com/auth/presentations"],
});

const slides = google.slides({ version: "v1", auth });
const presentationId = "1pdQFbo6pyXXAlsPyHw1Ho1KeewU791S554W6dQ19OHc";

async function run() {
  try {
    const presentation = await slides.presentations.get({ presentationId });
    const slideList = presentation.data.slides;
    const totalSlides = slideList.length;

    // Placeholders = last 2 slides
    const template1 = slideList[totalSlides - 2].objectId;
    const template2 = slideList[totalSlides - 1].objectId;

    // Step 1: Duplicate placeholders (they'll appear at the end by default)
    const duplicateRes = await slides.presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [
          { duplicateObject: { objectId: template1 } },
          { duplicateObject: { objectId: template2 } },
        ],
      },
    });

    // Get new slide IDs
    const newSlides = duplicateRes.data.replies.map(r => r.duplicateObject.objectId);

    // Step 2: Move them to the front (slide index 0 and 1)
    await slides.presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [
          {
            updateSlidesPosition: {
              slideObjectIds: [newSlides[0]],
              insertionIndex: 0,
            },
          },
          {
            updateSlidesPosition: {
              slideObjectIds: [newSlides[1]],
              insertionIndex: 1,
            },
          },
        ],
      },
    });

    // Example replacements
    const replacements = {
      "{{B6}}": "edwin (Aug–Oct 2024)",
      "{{C6}}": "edwin (Nov–Dec 2024)",
      "{{D6}}": "high-end brands (Jan–Dec 2024)",
      "{{B7}}": "80.18%", "{{C7}}": "82.40%", "{{D7}}": "72.96%",
      "{{B8}}": "3.12%", "{{C8}}": "3.48%", "{{D8}}": "2.86%",
      "{{B9}}": "98,477", "{{C9}}": "112,220", "{{D9}}": "4,017,030",
      "{{B10}}": "4.49%", "{{C10}}": "4.06%", "{{D10}}": "2.70%",
      "{{B11}}": "4,420", "{{C11}}": "4,554", "{{D11}}": "108,275",
      "{{B12}}": "21.14%", "{{C12}}": "22.29%", "{{D12}}": "13.54%",
      "{{B13}}": "17.12%", "{{C13}}": "20.81%", "{{D13}}": "12.66%",
      "{{B14}}": "38.26%", "{{C14}}": "43.10%", "{{D14}}": "26.19%",
      "{{B15}}": "996.60%", "{{C15}}": "1024.66%", "{{D15}}": "540.68%",
      "{{B16}}": "N/A", "{{C16}}": "N/A", "{{D16}}": "N/A",
      "{{B17}}": "N/A", "{{C17}}": "N/A", "{{D17}}": "N/A",
      "{{B18}}": "0.53%", "{{C18}}": "0.52%", "{{D18}}": "0.42%",
      "{{B19}}": "3,932,917", "{{C19}}": "3,908,977", "{{D19}}": "192,320,029",
      "{{B20}}": "3,153,459", "{{C20}}": "3,221,182", "{{D20}}": "140,316,101",
      "{{B21}}": "101,812", "{{C21}}": "111,343", "{{D21}}": "3,634,491",
      "{{B22}}": "681,219", "{{C22}}": "749,965", "{{D22}}": "N/A",
      "{{B23}}": "20,905", "{{C23}}": "20,427", "{{D23}}": "799,824",
      "{{B24}}": "3,578", "{{C24}}": "4,250", "{{D24}}": "101,223",
      "{{B25}}": "1,249,923", "{{C25}}": "1,128,798", "{{D25}}": "N/A",
      "{{B26}}": "857,508", "{{C26}}": "792,915", "{{D26}}": "N/A",
      "{{B27}}": "52,908", "{{C27}}": "56,287", "{{D27}}": "N/A",
      "{{B28}}": "46,246", "{{C28}}": "50,092", "{{D28}}": "N/A",
      "{{B29}}": "75,252", "{{C29}}": "77,003", "{{D29}}": "N/A",
      "{{B30}}": "12,759", "{{C30}}": "13,690", "{{D30}}": "N/A",
      "{{B31}}": "3,475", "{{C31}}": "3,732", "{{D31}}": "N/A",
      "{{B32}}": "2,375", "{{C32}}": "3,038", "{{D32}}": "N/A",
      "{{B33}}": "0.41%", "{{C33}}": "0.36%", "{{D33}}": "0.42%",
    };

    // Build replacement requests
    const requests = [];
    for (const [placeholder, value] of Object.entries(replacements)) {
      requests.push({
        replaceAllText: {
          containsText: { text: placeholder, matchCase: true },
          replaceText: value,
          pageObjectIds: newSlides, // ✅ only update new slides
        },
      });
    }

    await slides.presentations.batchUpdate({
      presentationId,
      requestBody: { requests },
    });

    console.log("✅ New slides created at the beginning!");
  } catch (err) {
    console.error("❌ Error:", err);
  }
}

run();