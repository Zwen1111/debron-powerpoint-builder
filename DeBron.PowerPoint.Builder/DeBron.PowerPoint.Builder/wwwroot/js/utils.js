window.downloadHelper = {
    downloadFileFromStream: async (fileName, contentStreamReference) => {
        const arrayBuffer = await contentStreamReference.arrayBuffer();
        const blob = new Blob([arrayBuffer]);
        const url = URL.createObjectURL(blob);
        const anchorElement = document.createElement('a');
        anchorElement.href = url;
        anchorElement.download = fileName ?? '';
        document.body.appendChild(anchorElement);
        anchorElement.click();
        anchorElement.remove();
        URL.revokeObjectURL(url);
    }
};

window.registerPasteHandler = function (dotNetHelper, elementId) {
    const element = document.getElementById(elementId);
    if (!element) return;

    element.addEventListener('paste', async function (event) {
        event.preventDefault(); // voorkom standaard plakken
        const pastedText = (event.clipboardData || window.clipboardData).getData('text');

        const transformed = await dotNetHelper.invokeMethodAsync('HandlePaste', pastedText);
        element.value = transformed;
    });
};

window.localStorageFunctions = {
    setItem: function (key, value) {
        localStorage.setItem(key, value);
    },
    getItem: function (key) {
        return localStorage.getItem(key);
    },
    removeItem: function (key) {
        localStorage.removeItem(key);
    }
};

