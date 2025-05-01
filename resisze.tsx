#target photoshop

function resizeAndCenterTopLayerToMatchBelow() {
    if (app.documents.length === 0) {
        alert("Không có file nào đang mở.");
        return;
    }

    var doc = app.activeDocument;
    var layers = doc.layers;

    if (layers.length < 2) {
        alert("Cần ít nhất 2 layer để chạy script.");
        return;
    }

    var topLayer = layers[0];
    var belowLayer = layers[1];

    var originalUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;

    // Lấy kích thước của layer dưới
    doc.activeLayer = belowLayer;
    var bBounds = belowLayer.bounds;
    var bX = (bBounds[0].as('px') + bBounds[2].as('px')) / 2;
    var bY = (bBounds[1].as('px') + bBounds[3].as('px')) / 2;
    var bW = bBounds[2].as('px') - bBounds[0].as('px');
    var bH = bBounds[3].as('px') - bBounds[1].as('px');

    // Resize layer trên cùng
    doc.activeLayer = topLayer;
    var tBounds = topLayer.bounds;
    var tW = tBounds[2].as('px') - tBounds[0].as('px');
    var tH = tBounds[3].as('px') - tBounds[1].as('px');

    var scaleX = (bW / tW) * 100;
    var scaleY = (bH / tH) * 100;

    topLayer.resize(scaleX, scaleY, AnchorPosition.MIDDLECENTER);

    // Canh giữa layer trên cùng theo layer dưới
    var newBounds = topLayer.bounds;
    var tCenterX = (newBounds[0].as('px') + newBounds[2].as('px')) / 2;
    var tCenterY = (newBounds[1].as('px') + newBounds[3].as('px')) / 2;

    var deltaX = bX - tCenterX;
    var deltaY = bY - tCenterY;

    topLayer.translate(deltaX, deltaY);

    app.preferences.rulerUnits = originalUnits;

    alert("Đã resize và canh giữa layer trên cùng.");
}

resizeAndCenterTopLayerToMatchBelow();
