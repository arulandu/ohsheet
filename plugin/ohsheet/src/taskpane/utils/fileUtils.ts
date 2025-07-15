export async function loadFilePath() {
  return new Promise((resolve) => {
    Office.context.document.getFilePropertiesAsync(null, (res) => {
      if (res && res.value && res.value.url) {
        let name = res.value.url.substr(res.value.url.lastIndexOf("\\") + 1);
        resolve(name);
      }
      resolve("");
    });
  });
} 