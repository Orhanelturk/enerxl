window.SearchTool = {
  init(map) {
    UI.register("search", "search-btn", "sub-bar-search");

    const input = document.getElementById("search-input");
    const autocomplete = new google.maps.places.Autocomplete(input);
    autocomplete.bindTo("bounds", map);

    autocomplete.addListener("place_changed", () => {
      const place = autocomplete.getPlace();
      if (!place.geometry) return;
      map.panTo(place.geometry.location);
      map.setZoom(18);
      UI.closeAll();
    });
  }
};
