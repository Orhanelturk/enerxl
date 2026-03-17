window.initMap = () => {
  const map = new google.maps.Map(document.getElementById("map"), {
    center: { lat: 38.40011471539006, lng: 34.97760350013513 },
    zoom: 18,
    mapTypeId: "satellite",
    disableDefaultUI: true
  });

  UI.register("system-inputs", "btn-system-inputs", "panel-system-inputs");
  UI.register("pv-design", "btn-pv-design", "panel-pv-design");
  UI.register("electrical-design", "btn-electrical-design", "panel-electrical-design");
  UI.register("generate-layout", "btn-generate-layout", "panel-generate-layout");

  UI.register("edit-tools", "btn-edit-tools", "panel-edit-tools");
  UI.register("layers", "btn-layers", "panel-layers");
  UI.register("system-summary", "btn-system-summary", "panel-system-summary");
  UI.register("boq", "btn-boq", "panel-boq");

  UI.register("library", "btn-library", "panel-library");
  UI.register("export", "btn-export", "panel-export");
  UI.register("save", "btn-save", "panel-save");

  UI.register("area", "area-btn", "sub-bar-area");

  UI.init();
  SearchTool.init(map);
  ProjectAreaTool.init(map);
};
