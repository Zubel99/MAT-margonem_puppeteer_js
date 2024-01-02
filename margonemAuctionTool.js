const puppeteer = require("puppeteer-extra");
const expect = require("chai").expect;
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const fs = require("fs");
const XLSX = require("xlsx");
const { spawn } = require("child_process"); // to run CNN captchaSolver.py
puppeteer.use(StealthPlugin());
require("dotenv").config();

const login = process.env.margonemLogin;
const password = process.env.margonemPassword;

const excelSavePath = "auctionsData";

//fix paths later
const pythonEnv =
  "D:\\Coding\\margonemAuctionToolBot\\CNN_python_env\\Scripts\\python.exe";
const scriptLoc =
  "D:\\Coding\\margonemAuctionToolBot\\captcha_CNN_py\\models\\research\\solveCaptcha.py";

(async () => {
  const browser = await puppeteer.launch({
    args: [
      `--proxy-server=http://${process.env.proxyAddress}:${process.env.proxyPort}`,
    ],
    headless: false,
    slowMo: 100,
    userDataDir: "./tmp", // try without later on and add handling for /intro
    devtools: false,
    timeout: 60000,
  });

  const delay = (milliseconds) =>
    new Promise((resolve) => setTimeout(resolve, milliseconds));
  function getRandomInt(min, max) {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min) + min);
  }
  const page = await browser.newPage();
  await page.setUserAgent(
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.125 Safari/537.36"
  );

  // **functions
  async function moveAndEnterAuctions() {
    await page.keyboard.press("KeyR");
    await delay(getRandomInt(500, 1000));
    await page.keyboard.press("KeyR");
    await delay(getRandomInt(500, 1000));
    await page.keyboard.press("KeyR");
    await delay(getRandomInt(500, 1000));
    await page.keyboard.press("1");
    await delay(getRandomInt(500, 1000));
    await page.keyboard.press("1");
    await delay(getRandomInt(500, 1000));
    await page.keyboard.press("1");
  }

  function jsScript() {
    return new Promise(async (resolve, reject) => {
      let serverName = location.href.substring(
        8,
        location.href.indexOf("margonem.pl") - 1
      );
      console.log("serverName", serverName);
      let tempDate = new Date();
      let date = Math.floor(tempDate.getTime() / 1000);
      let STOP_FETCH = false;

      function getRandomNumber(min, max) {
        return Math.random() * (max - min) + min;
      }

      const mousedownEvent = new MouseEvent("mousedown", {
        bubbles: true,
        cancelable: true,
        view: window,
      });
      const mouseupEvent = new MouseEvent("mouseup", {
        bubbles: true,
        cancelable: true,
        view: window,
      });
      //const mouseoutEvent = new MouseEvent('mouseout', {
      //   bubbles: true,
      //   cancelable: true,
      //   view: window
      //});
      const clickEvent = new MouseEvent("click", {
        bubbles: true,
        cancelable: true,
        view: window,
      });

      function setFilters() {
        let inputFields = document.querySelectorAll(
          "div.first-column-bar-search input"
        );
        inputFields.forEach((field) => {
          field.value = "";
        });
        let auctionType = document.querySelectorAll(
          "div.auction-type-wrapper div.bck-wrapper div.option"
        ); //0 = all, 1 = bid, 2 = buyout
        let itemRarity = document.querySelectorAll(
          "div.item-rarity-wrapper div.bck-wrapper div.option"
        ); //0 = all, 1 = normal, 2 = unique, 3 = heroic, 4 = legendary
        let itemProfession = document.querySelectorAll(
          "div.item-prof-wrapper div.bck-wrapper div.option"
        ); //0 = all, 1 = warrior, 2 = paladin, 3 = mage, 4 = hunter, 5 = blade dancer, 6 = tracker
        auctionType[2].dispatchEvent(clickEvent);
        itemRarity[3].dispatchEvent(clickEvent);
        itemProfession[0].dispatchEvent(clickEvent);
      }

      function getTargetCategories() {
        let allCategories = document.querySelectorAll(
          "div.one-category-auction[tip-id]"
        );
        let targetCategories = [
          ["Broń jednoręczna", 0],
          ["Broń dystansowa", 1],
          ["Broń dwuręczna", 2],
          ["Różdżki magiczne", 3],
          ["Orby magiczne", 4],
          ["Broń półtoraręczna", 5],
          ["Broń pomocnicza", 6],
          ["Strzały", 7],
          ["Zbroje", 8],
          ["Rękawice", 9],
          ["Buty", 10],
          ["Tarcze", 11],
          ["Hełmy", 12],
          ["Pierścienie", 13],
          ["Naszyjniki", 14],
          ["Neutralne", 17],
          ["Torby", 19],
        ];
        //let targetCategories = [['Różdżki magiczne', 3]]
        let filteredCategories = [];
        for (let i = 0; i < targetCategories.length; i++) {
          filteredCategories.push([
            targetCategories[i][0],
            allCategories[targetCategories[i][1]],
          ]);
        }
        console.log("filteredCategories", filteredCategories);
        return filteredCategories;
      }

      function loadAllSingleCategoryItems() {
        return new Promise((resolve) => {
          const downButton = document.querySelectorAll("div.arrow-down")[11];
          //const availableItemsAmount = parseInt(document.querySelector('div.amount-of-auction').innerText.replace(/\D/g, ''));
          let downBtnInterval;
          let checkAllItemsLoadedGuard;

          function handleClearInterval() {
            console.log("All items loaded. Clearing interval.");
            clearInterval(downBtnInterval);
            clearInterval(checkAllItemsLoadedGuard);
            resolve();
            console.log("clearInterval");
          }

          let numberOfAuctions = parseInt(
            document
              .querySelector("div.amount-of-auction")
              .innerText.replace(/\D/g, "")
          );
          const handleButtonClick = () => {
            console.log("Scrolling down");
            downButton.dispatchEvent(mousedownEvent);
            setTimeout(() => {
              downButton.dispatchEvent(mouseupEvent);
            }, getRandomNumber(4, 6));
          };

          const checkAllItemsLoaded = () => {
            const currentlyFetchedItems = document.querySelectorAll(
              "div.auction-window table.auction-table tr"
            );
            numberOfAuctions = parseInt(
              document
                .querySelector("div.amount-of-auction")
                .innerText.replace(/\D/g, "")
            );
            if (currentlyFetchedItems.length >= numberOfAuctions) {
              console.log("a");
              handleClearInterval();
            }
          };

          let currFetItems = [];
          checkAllItemsLoadedGuard = setInterval(() => {
            numberOfAuctions = parseInt(
              document
                .querySelector("div.amount-of-auction")
                .innerText.replace(/\D/g, "")
            );
            let buffer = document.querySelectorAll(
              "div.auction-window table.auction-table tr"
            );
            buffer.forEach((item) => {
              if (currFetItems.includes(item)) return;
              currFetItems.push(item);
            });
            console.log("auction count so far: ", currFetItems.length);
            console.log("numberOfAuctions: ", numberOfAuctions);
            if (currFetItems.length >= numberOfAuctions) {
              console.log("b");
              handleClearInterval();
            }
          }, 1000);

          downBtnInterval = setInterval(() => {
            if (STOP_FETCH) handleClearInterval();
            numberOfAuctions = parseInt(
              document
                .querySelector("div.amount-of-auction")
                .innerText.replace(/\D/g, "")
            );
            handleButtonClick();
            checkAllItemsLoaded();
          }, getRandomNumber(8, 10));
        });
      }

      function fetchAllSingleCategoryItems(categoryName) {
        let items = document.querySelectorAll(
          "div.auction-window table.auction-table tr"
        );
        let partialResultArray = [];

        items.forEach((item) => {
          let name = item.querySelector("td.item-name-td").innerText;
          let tempPrice = item.querySelector(
            "td.item-buy-now-td div.auction-cost-label"
          ).innerText;
          if (tempPrice.includes("SŁ")) return;
          let price;
          if (tempPrice.includes("k"))
            price = parseFloat(tempPrice.replace(/[^0-9.]/g, "")) * 1000;
          else if (tempPrice.includes("m"))
            price = parseFloat(tempPrice.replace(/[^0-9.]/g, "")) * 1000000;
          else if (tempPrice.includes("g"))
            price = parseFloat(tempPrice.replace(/[^0-9.]/g, "")) * 1000000000;
          else price = parseFloat(tempPrice);
          let level = parseInt(
            item.querySelector("td.item-level-td").innerText
          );
          let id = parseInt(
            item
              .querySelector("div.item[data-name]")
              .getAttribute("class")
              .replace(/\D/g, "")
          );
          let tempRarity = item
            .querySelector("div.item[data-name]")
            .getAttribute("data-item-type");
          let rarity;
          if (tempRarity == "t-uniupg") rarity = 0;
          else if (tempRarity == "t-her") rarity = 1;
          else if (tempRarity == "t-leg") rarity = 2;
          else {
            console.log("ignored rarity");
            return;
          }
          partialResultArray.push({
            categoryName: categoryName,
            itemName: name,
            itemRarity: rarity,
            itemPrice: price,
            itemLevel: level,
            itemId: id,
            timestamp: date,
          });
        });
        console.log(categoryName, ": ", partialResultArray);
        return partialResultArray;
      }

      function delay(ms) {
        return new Promise((resolve) => {
          setTimeout(() => {
            console.log("after delay");
            resolve();
          }, ms);
        });
      }

      // async function sendDataToServer(tableName, _data, _delay){
      //   const data = [tableName, _data];
      //   console.log('send object to local server: ', data)
      //
      //   fetch('http://localhost:3000/update-excel', {
      //     method: 'POST',
      //     headers: {
      //       'Content-Type': 'application/json'
      //     },
      //     body: JSON.stringify(data)
      //   })
      //       .then(response => response.json())
      //       .then(data => {
      //         console.log('Server response:', data);
      //       })
      //       .catch(error => {
      //         console.error('Error sending data:', error);
      //         let delay = (_delay + 1) * 1000;
      //         if (_delay < 10){
      //           console.log('retrying in: ', _delay, 's')
      //           setTimeout(() => sendDataToServer(tableName, _data, _delay + 1), delay);
      //         }
      //       });
      // }

      // ---------------

      function startLogic() {
        tempDate = new Date();
        date = Math.floor(tempDate.getTime() / 1000);
      }

      function fetchAuctionItems() {
        return new Promise(async (resolve, reject) => {
          startLogic();
          let resultArray = [];
          setFilters();
          await delay(500);
          let itemCategorise = getTargetCategories();
          for (const category of itemCategorise) {
            //category[0] = polish name, category[1] = html element
            if (STOP_FETCH) break;
            category[1].dispatchEvent(clickEvent);
            await delay(1000);

            await loadAllSingleCategoryItems();

            let fetchedData = await fetchAllSingleCategoryItems(category[0]);
            resultArray = resultArray.concat(fetchedData);
          }
          console.log("final array: ", resultArray);
          //if (resultArray.length > 0) {
          //  await saveDataToExcel([serverName + '_items', resultArray])
          //}
          STOP_FETCH = true;
          console.log("async test");
          resolve([serverName + "_items", resultArray]);
          console.log("fetch");
        });
      }

      let returnValue = await fetchAuctionItems();
      console.log("returnValue, ", returnValue);
      console.log("KONIEC");
      resolve(returnValue);
    });
  }

  function saveDataToExcel(data) {
    return new Promise((resolve, reject) => {
      console.log("[" + data[0] + "] Fetched " + data[1].length + " items");

      // Excel file path
      const excelFilePath = excelSavePath + "\\" + data[0] + ".xlsx";

      // Check if the Excel file doesnt exist
      if (!fs.existsSync(excelFilePath)) {
        // Create a new Excel workbook and worksheet
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet([]);

        // Append column names as the first row
        const columnNames = [
          "categoryName",
          "itemName",
          "itemRarity",
          "itemPrice",
          "itemLevel",
          "itemId",
          "timestamp",
        ];
        XLSX.utils.sheet_add_aoa(worksheet, [columnNames]);

        // Save the new Excel file
        XLSX.utils.book_append_sheet(workbook, worksheet, "DataSheet");
        XLSX.writeFile(workbook, excelFilePath);
      }

      // Check if the Excel file exists
      if (fs.existsSync(excelFilePath)) {
        try {
          // Load existing Excel file
          const workbook = XLSX.readFile(excelFilePath);
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];

          // Filter out existing items with the same itemId and itemPrice
          const existingData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
          });
          const newData = data[1].filter((newItem) => {
            return !existingData.some(
              (existingItem) =>
                existingItem[5] === newItem.itemId &&
                existingItem[3] === newItem.itemPrice
            );
          });

          // if (newData.length === 0) {
          //   console.log('[' + data[0] + '] No new items to add')
          //   resolve();
          // }

          // Append new data to the worksheet
          const rowsToAdd = newData.map((item) => [
            item.categoryName,
            item.itemName,
            item.itemRarity,
            item.itemPrice,
            item.itemLevel,
            item.itemId,
            item.timestamp,
          ]);
          const lastRow = XLSX.utils.decode_range(worksheet["!ref"]).e.r;
          XLSX.utils.sheet_add_aoa(worksheet, rowsToAdd, {
            origin: lastRow + 1,
          });

          // Save the updated Excel file
          XLSX.writeFile(workbook, excelFilePath);

          console.log(
            "[",
            displayCurrentTime(),
            " " +
              data[0] +
              "] Excel file updated successfully. Added " +
              newData.length +
              " new items"
          );
          resolve();
        } catch (error) {
          console.log(
            "[",
            displayCurrentTime(),
            " " +
              data[0] +
              "] Error updating Excel file: File is currently open or locked"
          );
          resolve();
        }
      }
    });
  }
  // **functions

  async function confirmAndEnter() {
    //console.log('CAE 1')
    const emailConfirm1 = await page.$(
      ".disabled-enter-game-info .close-game-info"
    );
    //console.log('CAE 2')
    const enterButton = await page.$(".box-enter .c-btn.enter-game");
    //console.log('CAE 3')
    if (!emailConfirm1) {
      //console.log('CAE 4')
      await enterButton.click();
    }

    const emailConfirm2 = await page.$(
      ".disabled-enter-game-info .close-game-info"
    );
    //console.log('CAE 5')
    if (emailConfirm2) {
      //console.log('CAE 6')
      await emailConfirm2.click();
    }

    //console.log('CAE 7')
    const enterButton2 = await page.$(".box-enter .c-btn.enter-game");
    if (enterButton2) {
      //console.log('CAE 8')
      await enterButton2.click();
    }
  }

  async function runPythonCaptchaSolver(imageBlob) {
    console.log("captcha initiated");
    return new Promise((resolve, reject) => {
      const childPython = spawn(pythonEnv, [scriptLoc, imageBlob]);
      let pythonOutput = "";

      childPython.stdout.on("data", (data) => {
        console.log(`stdout: ${data}`);
        pythonOutput = data.toString();
      });

      childPython.stderr.on("data", (data) => {
        console.log(`stderr: ${data}`);
      });

      childPython.on("close", (code) => {
        if (code === 0) {
          try {
            let importantData = pythonOutput.substring(
              pythonOutput.indexOf("SPLIT-FETCH") + 13
            );
            let parsedData = JSON.parse(importantData);
            console.log("parsed as array:", parsedData);
            resolve(parsedData);
          } catch (error) {
            console.error("Error parsing JSON:", error);
            reject(error);
          }
        } else {
          console.error(`Python script failed with exit code ${code}`);
          reject(new Error(`Python script failed with exit code ${code}`));
        }
      });
    });
  }

  async function handleCaptcha() {
    //first fetch captcha and instructions from website
    const preCaptchaButton = await page.$(
      ".captcha-pre-info .captcha-pre-info__button .button.green"
    );
    if (preCaptchaButton) {
      console.log("PRE CAPTCHA DETECTED");
      preCaptchaButton.click();
      console.log("a");
      await page.waitForSelector(".captcha-layer .border-window");
      console.log("b");
    }

    await delay(1000);
    let captchaWindow = await page.$(".captcha-layer .border-window");
    console.log("c");
    let captchaData;
    if (!captchaWindow) captchaData = null;
    else {
      console.log("d");
      captchaData = await page.evaluate(() => {
        console.log("checking for captcha");
        //let captchaWindow = document.querySelector('.captcha-layer .border-window');
        let menu = document.querySelector(
          ".captcha-layer .border-window .captcha__confirm"
        );

        let captchaImage = document.querySelector(
          ".captcha-layer .border-window .captcha__image"
        ).innerHTML;
        let srcData = captchaImage.substring(
          captchaImage.indexOf("base64,") + 7,
          captchaImage.indexOf("width") - 2
        );
        let instructions = document.querySelector(
          ".captcha-layer .border-window .captcha__question"
        ).innerText;
        // Create an object with the relevant data

        const captchaData = {
          imageBlob: "data:image/jpg;base64," + srcData,
          instructions: instructions,
        };
        console.log("imageBlob :", "data:image/jpg;base64,");
        console.log("raw data :", srcData);

        return [srcData, captchaData.instructions];
      });
    }
    //console.log('captcha: ', captchaData)

    //then check it with CNN
    console.log("e");
    if (captchaData == null) return 0;
    console.log("SEND THIS DATA TO MODEL: ", captchaData[0]);
    const imageClasses = await runPythonCaptchaSolver(captchaData[0]);
    const captchaButtons = await page.$$(
      ".captcha .captcha__buttons .button.green"
    ); // order of buttons ?? moze trzeba kliknac na background ? albo label?
    console.log(captchaButtons);
    for (let i = 0; i < captchaButtons.length; i++) {
      if (imageClasses[i] == "Upside Down") {
        await delay(getRandomInt(800, 1500));
        captchaButtons[i].click();
      }
    }
    //click buttons
    console.log("clicked all buttons");
    await delay(getRandomInt(300, 700));
    const confirmButton = await page.$(".captcha__confirm .button.green");
    console.log("click confirm button");
    confirmButton.click();
    return 1;
  }

  async function checkLogin() {
    const loginForm = await page.$$(
      ".login-form.info-container .form-group"
    )[0];
    if (loginForm) {
      console.log("logging in account");
      await page.$eval(
        ".login-form.info-container .form-group #login-input",
        (el, login) => {
          el.value = login;
        },
        login
      );
      await page.$eval(
        ".login-form.info-container .form-group #login-password",
        (el, password) => {
          el.value = password;
        },
        password
      );
      await page.click("#js-login-btn");
      console.log("logged in");
    }
  }
  function displayCurrentTime() {
    var today = new Date();
    var time =
      today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    return time;
  }

  // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN *** // MAIN ***

  await page.goto("https://margonem.pl/", { waitUntil: "load" });
  await checkLogin();

  await page.waitForSelector(".select-char"); // front
  await page.click(".select-char");
  await page.waitForSelector(".charlist"); // back
  const playerChars = await page.$$(".charlist .charc"); // back
  await playerChars[0].click();
  //for (const char of playerChars) {
  for (let i = 0; i < playerChars.length; i++) {
    console.log("waiting 1");
    await page.waitForSelector(".select-char");
    console.log("clicking");
    await page.click(".select-char");
    console.log("waiting 2");
    await page.waitForSelector(".charlist");
    console.log("get characters");
    const temp = await page.$$(".charlist .charc");
    console.log("click character");
    await temp[i].click();
    console.log("CAE");
    await confirmAndEnter();

    await page.waitForSelector("body[style]");
    console.log("body");

    let captchaCondition = 1;
    while (captchaCondition == 1) {
      captchaCondition = await handleCaptcha();
    }
    console.log("after captcha checkpoint 1");
    (await page.waitForSelector(".chat-size-1")) ||
      (await page.waitForSelector(".chat-size-0"));
    await moveAndEnterAuctions();
    await page.waitForSelector("div.content div.auction-window");
    const excelData = await page.evaluate(jsScript);
    await saveDataToExcel(excelData);
    //await handleCaptcha();

    captchaCondition = 1;
    while (captchaCondition == 1) {
      captchaCondition = await handleCaptcha();
    }
    console.log("after captcha checkpoint 2");
    await Promise.all([
      page.waitForNavigation({ waitUntil: "load" }),
      await page.goBack(),
    ]);
  }
  await browser.close();
})();
