// Отключаем диалоговые окна и цветовые профили
app.displayDialogs = DialogModes.NO;
app.preferences.askMissingProfile = false;
app.preferences.askProfileMismatch = false;
app.preferences.askPastedProfile = false;

// Лог-файл
var logFile = new File("D:/PyCharm 2024.2.3/PycharmProjects/fillpercent/log.txt");
logFile.open("w");
logFile.writeln("✅ Скрипт `blackfill.jsx` запущен!");

// Функция логирования
function logMessage(message) {
    logFile.writeln(message);
    logFile.flush();
}

// Получаем путь из аргументов
var folderPath = null;
if (arguments.length > 0 && arguments[0] && arguments[0] !== "undefined") {
    folderPath = decodeURI(arguments[0]);  // Декодируем путь
    logMessage("📂 Получена папка: " + folderPath);
} else {
    logMessage("❌ Photoshop НЕ получил путь! Используем автопоиск.");
    folderPath = "D:/PyCharm 2024.2.3/PycharmProjects/fillpercent/raw";
}

// Проверяем существование папки
var folder = new Folder(folderPath);
if (!folder.exists) {
    logMessage("❌ Ошибка: Папка не найдена!");
    logFile.close();
    return;
}

// Получаем ВСЕ PDF-файлы в папке
var files = folder.getFiles(function(file) {
    return file instanceof File && file.name.match(/\.pdf$/i);
});

if (files.length === 0) {
    logMessage("❌ В папке нет PDF-файлов!");
    logFile.close();
    return;
}

// Логируем найденные файлы
for (var i = 0; i < files.length; i++) {
    logMessage("📄 Найден PDF: " + files[i].fsName);
}

// Обрабатываем файлы ПО ОДНОМУ
for (var i = 0; i < files.length; i++) {
    processPDF(files[i]);
}

// Основная обработка PDF
function processPDF(fileToOpen) {
    if (!fileToOpen || !fileToOpen.exists) {
        logMessage("❌ Ошибка: PDF не найден!");
        return;
    }

    try {
        logMessage("📄 Открываем PDF: " + fileToOpen.fsName);

        var pdfOptions = new PDFOpenOptions();
        pdfOptions.page = 1;
        pdfOptions.openPages = 1;
        pdfOptions.mode = OpenDocumentMode.RGB;
        pdfOptions.resolution = 300;
        pdfOptions.constrainProportions = true;
        pdfOptions.antiAlias = true;
        pdfOptions.bitsPerChannel = BitsPerChannelType.EIGHT;

        app.open(fileToOpen, pdfOptions);

        if (!waitForDocument(15)) {
            logMessage("❌ Ошибка: документ не открылся!");
            return;
        }

        var doc = app.activeDocument;
        deleteHiddenLayers();
        doc.mergeVisibleLayers();
        selectLayerTransparency();

        var blackColor = new SolidColor();
        blackColor.rgb.red = 0;
        blackColor.rgb.green = 0;
        blackColor.rgb.blue = 0;
        doc.selection.fill(blackColor, ColorBlendMode.NORMAL, 100, false);
        doc.selection.deselect();

        var originalName = doc.name.replace(/\.[^\.]+$/, '');
        var pdfPath = "D:\\PyCharm 2024.2.3\\PycharmProjects\\fillpercent\\Prints\\" + originalName + ".pdf";
        var pdfFile = new File(pdfPath);

        var pdfSaveOptions = new PDFSaveOptions();
        pdfSaveOptions.embedColorProfile = true;
        pdfSaveOptions.preserveEditing = false;
        pdfSaveOptions.optimization = true;
        pdfSaveOptions.viewAfterSaving = false;

        try {
            doc.saved = true;
            doc.saveAs(pdfFile, pdfSaveOptions, true, Extension.LOWERCASE);
            logMessage("✅ Документ успешно сохранён: " + pdfPath);
        } catch (e) {
            logMessage("❌ Ошибка при сохранении PDF: " + e);
        }
    } catch (e) {
        logMessage("❌ Ошибка при открытии PDF: " + e);
    }
}

// Функция ожидания загрузки документа
function waitForDocument(timeoutInSeconds) {
    var startTime = new Date().getTime();
    while (app.documents.length === 0) {
        if ((new Date().getTime() - startTime) > timeoutInSeconds * 1000) {
            logMessage("❌ Файл не открылся в Photoshop!");
            return false;
        }
        $.sleep(500);
    }
    return true;
}

// Закрываем лог
logFile.close();
