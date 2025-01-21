/*
 * (c) Microsoft Corporation. Licensed under the MIT license.
 *
 * Office.js の型定義を用いて PowerPoint アドインを TypeScript で実装したサンプル。
 * base64Image.ts から Base64 文字列を import して、setSelectedDataAsync で画像を挿入。
 */

import { base64Image } from "../../base64Image";
// ↑ base64Image を定義しているファイルを正しいパスで import してください。
// 例: export const base64Image = "data:image/png;base64, ...";
import pptxgen from "pptxgenjs";

// Office.js の型定義を使いたい場合は、@types/office-js をインストールし、
// ついでに下記の import を入れると VSCode などで IntelliSense が利きやすくなります。
// import "office-js";

/* global document, Office, PowerPoint, Rollbar */

/**
 * Office が準備できたタイミングで呼ばれる。
 * PowerPoint でホストされる場合のみ UI を表示し、ボタンのイベントを設定。
 */
Office.onReady((info: Office.HostReadyInfo) => {
  if (info.host === Office.HostType.PowerPoint) {
    const sideloadMsg = document.getElementById("sideload-msg") as HTMLElement | null;
    const appBody = document.getElementById("app-body") as HTMLElement | null;

    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    if (appBody) {
      appBody.style.display = "flex";
    }

    // 「insert-image」ボタンにクリックイベントを設定
    const insertImageButton = document.getElementById("insert-image") as HTMLButtonElement | null;
    if (insertImageButton) {
      insertImageButton.onclick = () => clearMessage(insertImage);
    }

    const insertTextButton = document.getElementById("insert-text") as HTMLButtonElement | null;
    if (insertTextButton) {
      insertTextButton.onclick = () => clearMessage(insertText);
    }

    const getSlideMetadataButton = document.getElementById("get-slide-metadata") as HTMLButtonElement | null;
    if (getSlideMetadataButton) {
      getSlideMetadataButton.onclick = () => tryCatch(getSlideMetadata);
    }

    const addSlidesButton = document.getElementById("add-slides") as HTMLButtonElement | null;
    if (addSlidesButton) {
      addSlidesButton.onclick = () => tryCatch(addSlides);
    }
    const goToFirstSlideButton = document.getElementById("go-to-first-slide") as HTMLButtonElement | null;
    if (goToFirstSlideButton) {
      goToFirstSlideButton.onclick = () => clearMessage(goToFirstSlide);
    }
    const goToNextSlideButton = document.getElementById("go-to-next-slide") as HTMLButtonElement | null;
    if (goToNextSlideButton) {
      goToNextSlideButton.onclick = () => clearMessage(goToNextSlide);
    }
    const goToPreviousSlideButton = document.getElementById("go-to-previous-slide") as HTMLButtonElement | null;
    if (goToPreviousSlideButton) {
      goToPreviousSlideButton.onclick = () => clearMessage(goToPreviousSlide);
    }
    const goToLastSlideButton = document.getElementById("go-to-last-slide") as HTMLButtonElement | null;
    if (goToLastSlideButton) {
      goToLastSlideButton.onclick = () => clearMessage(goToLastSlide);
    }

    //「Create PPTX from Message」ボタンにクリックイベントを設定
    const createPptButton = document.getElementById("create-ppt-from-message") as HTMLButtonElement | null;
    if (createPptButton) {
      createPptButton.onclick = () => tryCatch(addSlideFromPpt);
    }
  }
});

/**
 * 画像を挿入する非同期処理。Promise<void> を返すように実装。
 */
function insertImage(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      base64Image,
      { coercionType: Office.CoercionType.Image },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
          reject(asyncResult.error);
        } else {
          resolve();
        }
      }
    );
  });
}

async function insertText(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      "Hello World!",
      { coercionType: Office.CoercionType.Text },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
          reject(asyncResult.error);
        } else {
          resolve();
        }
      }
    );
  });
}

async function getSlideMetadata(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.SlideRange,
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
          reject(asyncResult.error);
        } else {
          setMessage("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
          resolve();
        }
      }
    );
  });
}

/**
 * 新しいスライドを 2 枚追加し、最後のスライドへ移動する。
 * PowerPoint.run(...) を使うには、@types/office-js-preview 等が必要な場合があります。
 */
async function addSlides(): Promise<void> {
  await PowerPoint.run(async (context: any) => {
    context.presentation.slides.add();
    context.presentation.slides.add();
    await context.sync();

    await goToLastSlide(); // 追加後に最後のスライドに移動
    setMessage("Success: Slides added.");
  });
}

/**
 * 最初のスライドへ移動
 */
function goToFirstSlide(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
        reject(asyncResult.error);
      } else {
        resolve();
      }
    });
  });
}

/**
 * 最後のスライドへ移動
 */
function goToLastSlide(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
        reject(asyncResult.error);
      } else {
        resolve();
      }
    });
  });
}

/**
 * 前のスライドへ移動
 */
function goToPreviousSlide(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
        reject(asyncResult.error);
      } else {
        resolve();
      }
    });
  });
}

/**
 * 次のスライドへ移動
 */
function goToNextSlide(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
        reject(asyncResult.error);
      } else {
        resolve();
      }
    });
  });
}

/**
 * PptxGenJS を使って、現在のスライドを base64 に変換し、追加する。
 *
 */
async function addSlideFromPpt(): Promise<void> {
  const pptx = new pptxgen();
  const slide1 = pptx.addSlide();
  const buttonElement = document.getElementById("create-ppt-from-message") as HTMLElement | null;
  slide1.addText(buttonElement.innerText, { x: 1, y: 1, w: 5, h: 1, fontSize: 18, color: "FF0000" });
  // 画像を追加する場合は、以下のようにする
  // slide1.addImage({ path: "https://upload.wikimedia.org/wikipedia/en/a/a9/Example.jpg", x: 0.5, y: 0.5, w: 8, h: 6 });
  const base64 = await pptx.write("base64");

  await PowerPoint.run(async (context) => {
    context.presentation.insertSlidesFromBase64(base64, {
      formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
    });
    await context.sync();
  });
}

/**
 * メッセージをクリアした後に、コールバック関数を実行する。
 */
async function clearMessage(callback: () => Promise<void>): Promise<void> {
  const messageElement = document.getElementById("message") as HTMLElement | null;
  if (messageElement) {
    messageElement.innerText = "";
  }
  await callback(); // callbackがPromise<void>なら、ここで非同期的に完了を待てる
}

/**
 * メッセージを表示する。
 */
async function setMessage(message: string): Promise<void> {
  const messageElement = document.getElementById("message") as HTMLElement | null;
  if (messageElement) {
    messageElement.innerText = message;
  }
}

/**
 * 例外処理付きの実行ヘルパー関数 (必要に応じて呼び出す)
 */
async function tryCatch(callback: () => Promise<void>): Promise<void> {
  try {
    const messageElement = document.getElementById("message") as HTMLElement | null;
    if (messageElement) {
      messageElement.innerText = "";
    }
    await callback();
  } catch (error: unknown) {
    setMessage("Error: " + String(error));
  }
}
