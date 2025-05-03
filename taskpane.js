Office.onReady(() => {
  document.getElementById("shapefileInput").addEventListener("change", handleShapefileUpload);
});

async function handleShapefileUpload(event) {
  const files = event.target.files;
  const shpFile = Array.from(files).find(f => f.name.endsWith(".shp"));
  const dbfFile = Array.from(files).find(f => f.name.endsWith(".dbf"));
  if (!shpFile || !dbfFile) {
    alert("Por favor, selecione os arquivos .shp e .dbf.");
    return;
  }

  const shpArrayBuffer = await shpFile.arrayBuffer();
  const dbfArrayBuffer = await dbfFile.arrayBuffer();

  const source = await shapefile.open(shpArrayBuffer, dbfArrayBuffer);
  const features = [];
  while (true) {
    const result = await source.read();
    if (result.done) break;
    features.push(result.value);
  }

  // Enviar atributos para Excel
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headers = Object.keys(features[0].properties);
    const data = features.map(f => headers.map(h => f.properties[h]));
    sheet.getRange("A1").values = [headers, ...data];
    await context.sync();
  });

  // Mostrar geometrias no mapa
  const geojson = {
    type: "FeatureCollection",
    features: features.map(f => ({ type: "Feature", geometry: f.geometry }))
  };

  const map = L.map("map").setView([0, 0], 2);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors"
  }).addTo(map);

  L.geoJSON(geojson).addTo(map);
}