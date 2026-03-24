app.displayDialogs = DialogModes.NO; // Отключаем диалоговые окна
app.preferences.askMissingProfile = false;
app.preferences.askProfileMismatch = false;
app.preferences.askPastedProfile = false;

var outputFolder = new Folder("D:/PyCharm 2024.2.3/PycharmProjects/fillpercent/Prints");
if (!outputFolder.exists) {
    outputFolder.create();
}

// 🔹 Запрашиваем режим работы
var mode = confirm("Вы хотите обработать ВСЕ PDF в папке? (OK - выбрать папку, Отмена - выбрать файлы вручную)");

var filesToProcess = [];

if (mode) {
    var folder = Folder.selectDialog("Выберите папку с PDF файлами");
    if (folder) {
        var files = folder.getFiles("*.pdf");
        if (files.length > 0) {
            filesToProcess = files;
        } else {
            alert("❌ В папке нет PDF-файлов!");
        }
    } else {
        alert("❌ Папка не выбрана.");
    }
} else {
    filesToProcess = File.openDialog("Выберите PDF-файлы для обработки", "*.pdf", true);
}

if (filesToProcess.length === 0) {
    alert("❌ Не выбраны файлы для обработки!");
} else {
    for (var i = 0; i < filesToProcess.length; i++) {
        processPDF(filesToProcess[i]);
    }
    alert("✅ Все файлы обработаны и сохранены в: " + outputFolder.fsName);
}

// 🔹 Функция обработки PDF
function processPDF(fileToOpen) {
    if (fileToOpen && fileToOpen.exists) {
        try {
            var pdfOptions = new PDFOpenOptions();
            pdfOptions.page = 1;
            pdfOptions.openPages = 1;
            pdfOptions.mode = OpenDocumentMode.RGB;
            pdfOptions.resolution = 300;
            pdfOptions.constrainProportions = true;
            pdfOptions.antiAlias = true;
            pdfOptions.bitsPerChannel = BitsPerChannelType.EIGHT;

            app.open(fileToOpen, pdfOptions);

            if (app.documents.length === 0) {
                alert("❌ Ошибка: документ " + fileToOpen.name + " не открылся!");
                return;
            }

            var doc = app.activeDocument;
            deleteHiddenLayers();

            if (doc.artLayers.length > 1 || doc.layerSets.length > 0) {
                try {
                    doc.mergeVisibleLayers();
                } catch (e) {
                    alert("⚠ Merge Visible не выполнен: " + e);
                }
            }

            // 🔹 Проверяем, является ли слой Smart Object, и растрируем его
            rasterizeIfSmartObject();

            selectLayerTransparency();

            var blackColor = new SolidColor();
            blackColor.rgb.red = 0;
            blackColor.rgb.green = 0;
            blackColor.rgb.blue = 0;
            doc.selection.fill(blackColor, ColorBlendMode.NORMAL, 100, false);
            doc.selection.deselect();

            var outputPath = outputFolder.fsName + "/processed_" + fileToOpen.name;
            var pdfFile = new File(outputPath);

            var pdfSaveOptions = new PDFSaveOptions();
            pdfSaveOptions.embedColorProfile = true;
            pdfSaveOptions.preserveEditing = false;
            pdfSaveOptions.optimization = true;
            pdfSaveOptions.viewAfterSaving = false;

            doc.saveAs(pdfFile, pdfSaveOptions, true, Extension.LOWERCASE);
            doc.close(SaveOptions.DONOTSAVECHANGES);
            
        } catch (e) {
            alert("❌ Ошибка при обработке " + fileToOpen.name + ": " + e);
        }
    }
}

// 🔹 Функция удаления скрытых слоев
function deleteHiddenLayers() {
    var hiddenLayers = [];
    var layers = app.activeDocument.artLayers;
    for (var i = 0; i < layers.length; i++) {
        if (!layers[i].visible) {
            hiddenLayers.push(layers[i]);
        }
    }
    var groups = app.activeDocument.layerSets;
    for (var j = 0; j < groups.length; j++) {
        if (!groups[j].visible) {
            hiddenLayers.push(groups[j]);
        }
    }
    for (var k = hiddenLayers.length - 1; k >= 0; k--) {
        hiddenLayers[k].remove();
    }
}

// 🔹 Функция проверки и растрирования Smart Object
function rasterizeIfSmartObject() {
    var doc = app.activeDocument;
    var activeLayer = doc.activeLayer;
    
    if (activeLayer.kind == LayerKind.SMARTOBJECT) {
        try {
            activeLayer.rasterize(RasterizeType.ENTIRELAYER);
            alert("✅ Smart Object растрирован!");
        } catch (e) {
            alert("❌ Ошибка при растрировании: " + e);
        }
    }
}

// 🔹 Функция выбора прозрачных пикселей
function selectLayerTransparency() {
    var idsetd = charIDToTypeID("setd");
    var desc = new ActionDescriptor();
    var idnull = charIDToTypeID("null");
    var ref = new ActionReference();
    var idChnl = charIDToTypeID("Chnl");
    var idfsel = charIDToTypeID("fsel");
    ref.putProperty(idChnl, idfsel);
    desc.putReference(idnull, ref);
    var idT = charIDToTypeID("T   ");
    var ref2 = new ActionReference();
    var idChnl = charIDToTypeID("Chnl");
    var idTrsp = stringIDToTypeID("transparencyEnum");
    ref2.putEnumerated(idChnl, idChnl, idTrsp);
    desc.putReference(idT, ref2);
    executeAction(idsetd, desc, DialogModes.NO);
}
