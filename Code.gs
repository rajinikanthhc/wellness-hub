const SPREADSHEET_ID =
  "1shpwWTuTBC3waCjWLkANZOsP7mSmv1cK6Mt5J7Nz_HM";

const DELETE_PASSCODE = "12345";


/* =================================
   GITHUB CONFIG
   SAME METHOD AS VISITING CARDS
================================= */

const GITHUB_OWNER = "rajinikanthhc";
const GITHUB_REPO = "images";
const GITHUB_BRANCH = "main";


/* =================================
   WEB APP
================================= */

function doGet() {

  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("Wellness Hub");

}


function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* =================================
   GET SHEET
================================= */

function getSheet_(type) {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  const sheetName =
    type === "yoga"
      ? "Yoga"
      : "Exercise";


  const sheet =
    ss.getSheetByName(sheetName);


  if (!sheet) {

    throw new Error(
      "Sheet not found: " +
      sheetName
    );

  }


  return sheet;

}


/* =================================
   GET YOGA
================================= */

function getYoga() {

  return getItems_("yoga");

}


/* =================================
   GET EXERCISES
================================= */

function getExercises() {

  return getItems_("exercise");

}


/* =================================
   GET ITEMS
================================= */

function getItems_(type) {

  const sheet =
    getSheet_(type);


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return [];

  }


  /* ================================
     READ:
     ID
     Name
     Duration
     Image
     Notes
  ================================= */

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        5
      )
      .getValues();


  return data

    .filter(function(row) {

      return row[0];

    })

    .map(function(row) {

      const id =
        String(row[0] || "");


      const name =
        String(row[1] || "");


      const duration =
        String(row[2] || "");


      const imageName =
        String(row[3] || "")
          .trim();


      const notes =
        String(row[4] || "")
          .trim();


      let imageUrl = "";


      if (imageName) {

        imageUrl =
          "https://raw.githubusercontent.com/" +
          GITHUB_OWNER +
          "/" +
          GITHUB_REPO +
          "/" +
          GITHUB_BRANCH +
          "/" +
          getGitHubFolder_(type) +
          "/" +
          encodeURIComponent(imageName);

      }


      return {

        id: id,

        name: name,

        duration: duration,

        image: imageUrl,

        imageName: imageName,

        notes: notes

      };

    });

}

/* =================================
   GITHUB FOLDER
================================= */

function getGitHubFolder_(type) {

  return type === "yoga"
    ? "yoga"
    : "exercise";

}


/* =================================
   GET IMAGE URL
================================= */

function getImageUrl_(
  type,
  value
) {

  if (!value) {

    return "";

  }


  /*
     If old Sheet value is already
     a full URL, keep it.
  */

  if (
    value.indexOf(
      "http://"
    ) === 0 ||

    value.indexOf(
      "https://"
    ) === 0
  ) {

    return value;

  }


  const folder =
    type === "yoga"
      ? "yoga"
      : "exercise";


  return (
    "https://raw.githubusercontent.com/" +
    GITHUB_OWNER +
    "/" +
    GITHUB_REPO +
    "/" +
    GITHUB_BRANCH +
    "/" +
    folder +
    "/" +
    encodeURIComponent(value)
  );

}


/* =================================
   ADD ITEM
================================= */

function addItem(
  type,
  item
) {

  if (
    !item ||
    !item.name
  ) {

    throw new Error(
      "Name is required."
    );

  }


  const sheet =
    getSheet_(type);


  /*
     IMPORTANT:
     Lowest available ID
  */

  const id =
    generateNextId_(
      sheet,
      type
    );


  /*
     Only filename goes to Sheet
  */

  const imageName =
    extractImageFileName_(
      item.image || ""
    );


  sheet.appendRow([
  id,
  item.name,
  item.duration || "",
  imageName,
  item.notes || ""
]);


  return id;

}


/* =================================
   LOWEST AVAILABLE ID
================================= */

function generateNextId_(
  sheet,
  type
) {

  const prefix =
    type === "yoga"
      ? "Y"
      : "E";


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return prefix + "001";

  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues()
      .flat();


  const usedNumbers =
    new Set();


  ids.forEach(function(id) {

    const match =
      String(id)
        .trim()
        .match(
          new RegExp(
            "^" +
            prefix +
            "(\\d+)$",
            "i"
          )
        );


    if (match) {

      usedNumbers.add(
        Number(match[1])
      );

    }

  });


  let number = 1;


  while (
    usedNumbers.has(number)
  ) {

    number++;

  }


  return (
    prefix +
    String(number)
      .padStart(3, "0")
  );

}


/* =================================
   SAVE / EDIT ITEM
================================= */

function saveItem(
  type,
  item
) {

  if (
    !item ||
    !item.id
  ) {

    throw new Error(
      "Invalid record."
    );

  }


  const sheet =
    getSheet_(type);


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      "No records found."
    );

  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues()
      .flat();


  const index =
    ids.findIndex(function(id) {

      return (
        String(id) ===
        String(item.id)
      );

    });


  if (index === -1) {

    throw new Error(
      "Record not found: " +
      item.id
    );

  }


  const rowNumber =
    index + 2;


  /*
     Convert URL to filename
     if old record has URL.
  */

  const imageName =
    extractImageFileName_(
      item.image || ""
    );


  sheet
    .getRange(
      rowNumber,
      1,
      1,
      4
    )
    .setValues([[

      item.id,

      item.name || "",

      item.duration || "",

      imageName

    ]]);


  return true;

}


/* =================================
   DELETE ITEM
================================= */

function deleteItem(
  type,
  id,
  passcode
) {

  if (
    String(passcode) !==
    DELETE_PASSCODE
  ) {

    throw new Error(
      "Incorrect passcode."
    );

  }


  const sheet =
    getSheet_(type);


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      "No records found."
    );

  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues()
      .flat();


  const index =
    ids.findIndex(function(value) {

      return (
        String(value) ===
        String(id)
      );

    });


  if (index === -1) {

    throw new Error(
      "Record not found."
    );

  }


  sheet.deleteRow(
    index + 2
  );


  return true;

}


/* =================================
   GET ITEM
================================= */

function getItem(
  type,
  id
) {

  const items =
    getItems_(type);


  return (
    items.find(function(item) {

      return (
        String(item.id) ===
        String(id)
      );

    }) || null
  );

}


/* =================================
   IMAGE FILE NAME
================================= */

function createImageFileName_(
  type,
  name,
  extension
) {

  const folder =
    type === "yoga"
      ? "yoga"
      : "exercise";


  const slug =
    String(
      name || "image"
    )

      .toLowerCase()

      .trim()

      .replace(
        /[^a-z0-9]+/g,
        "-"
      )

      .replace(
        /^-+|-+$/g,
        ""
      );


  return (
    folder +
    "/" +
    slug +
    extension
  );

}


/* =================================
   EXTRACT FILENAME
================================= */

function extractImageFileName_(
  value
) {

  if (!value) {

    return "";

  }


  let filename =
    String(value)
      .trim();


  /*
     If it is a URL,
     take only the last part.
  */

  if (
    filename.indexOf(
      "http://"
    ) === 0 ||

    filename.indexOf(
      "https://"
    ) === 0
  ) {

    filename =
      filename
        .split("/")
        .pop();

  }


  /*
     Remove query string
  */

  filename =
    filename.split("?")[0];


  /*
     Decode URL characters
  */

  try {

    filename =
      decodeURIComponent(
        filename
      );

  } catch (e) {

  }


  return filename;

}


/* =================================
   GITHUB IMAGE UPLOAD
================================= */

function uploadImageToGitHub(
  type,
  name,
  base64Data,
  mimeType
) {

  if (!base64Data) {

    throw new Error(
      "No image received."
    );

  }


  /*
     ONLY TOKEN COMES FROM
     SCRIPT PROPERTIES
  */

  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "GITHUB_TOKEN"
      );


  if (!token) {

    throw new Error(
      "GITHUB_TOKEN not found in Script Properties."
    );

  }


  const extension =
    getExtensionFromMime_(
      mimeType
    );


  const fullPath =
    createImageFileName_(
      type,
      name,
      extension
    );


  /*
     Correct GitHub API path
  */

  const apiPath =
    fullPath
      .split("/")
      .map(function(part) {

        return encodeURIComponent(
          part
        );

      })
      .join("/");


  const url =
    "https://api.github.com/repos/" +
    GITHUB_OWNER +
    "/" +
    GITHUB_REPO +
    "/contents/" +
    apiPath;


  const cleanBase64 =
    String(base64Data)
      .replace(
        /^data:[^;]+;base64,/,
        ""
      );


  /*
     Check if file already exists
  */

  let sha = null;


  const checkResponse =
    UrlFetchApp.fetch(

      url +
      "?ref=" +
      encodeURIComponent(
        GITHUB_BRANCH
      ),

      {

        method: "get",

        headers: {

          Authorization:
            "Bearer " +
            token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        muteHttpExceptions:
          true

      }

    );


  if (
    checkResponse
      .getResponseCode() === 200
  ) {

    const existing =
      JSON.parse(
        checkResponse
          .getContentText()
      );


    sha =
      existing.sha;

  }


  const payload = {

    message:
      "Upload " +
      type +
      " image - " +
      name,

    content:
      cleanBase64,

    branch:
      GITHUB_BRANCH

  };


  if (sha) {

    payload.sha =
      sha;

  }


  const response =
    UrlFetchApp.fetch(

      url,

      {

        method: "put",

        contentType:
          "application/json",

        headers: {

          Authorization:
            "Bearer " +
            token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions:
          true

      }

    );


  const code =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  if (
    code < 200 ||
    code >= 300
  ) {

    throw new Error(
      "GitHub upload failed: " +
      responseText
    );

  }


  const filename =
    fullPath
      .split("/")
      .pop();


  const imageUrl =
    getImageUrl_(
      type,
      filename
    );


  /*
     Return both.
     Sheet will use filename.
     Frontend can use URL if needed.
  */

  return {

    fileName:
      filename,

    url:
      imageUrl

  };

}


/* =================================
   IMAGE EXTENSION
================================= */

function getExtensionFromMime_(
  mimeType
) {

  const map = {

    "image/jpeg": ".jpg",

    "image/jpg": ".jpg",

    "image/png": ".png",

    "image/webp": ".webp",

    "image/gif": ".gif",

    "image/bmp": ".bmp",

    "image/tiff": ".tiff"

  };


  return (
    map[mimeType] ||
    ".jpg"
  );

}

/* =================================
   OPEN GOOGLE SHEET
================================= */

function getSpreadsheetUrl() {

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getUrl();

}