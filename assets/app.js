(function () {
  var PIN_KEY = "faf-ps-pins";

  function sortKey(name) {
    return (name || "").replace(/^[^a-zA-Z0-9]+/, "").toLowerCase();
  }

  function letterFor(name) {
    var k = sortKey(name);
    var c = k.charAt(0).toUpperCase();
    return /[A-Z]/.test(c) ? c : "#";
  }

  function norm(s) {
    return (s || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "");
  }

  function formatDate(iso) {
    if (!iso) return "";
    var p = iso.split("-");
    if (p.length !== 3) return iso;
    var d = new Date(+p[0], +p[1] - 1, +p[2]);
    return (
      "Updated " +
      d.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
        year: "numeric",
      })
    );
  }

  function yyyymmdd(iso) {
    if (!iso) return "";
    return iso.replace(/-/g, "");
  }

  function builderSearchBlob(b) {
    return norm(b.name) + " " + norm((b.pdf || "").replace(/^.*\//, ""));
  }

  function loadPins() {
    try {
      return JSON.parse(localStorage.getItem(PIN_KEY) || "[]") || [];
    } catch (e) {
      return [];
    }
  }

  function savePins(pins) {
    try {
      localStorage.setItem(PIN_KEY, JSON.stringify(pins));
    } catch (e) {
      /* ignore */
    }
  }

  function stableId(b) {
    return norm(b.name);
  }

  function renderBuilderItem(b) {
    var id = stableId(b);
    var li = document.createElement("li");
    li.className = "ps-item";
    li.dataset.id = id;
    li.dataset.s = builderSearchBlob(b);
    li.dataset.u = yyyymmdd(b.updated);
    li.dataset.t = "";

    var pdfUrl = b.pdf;
    var pin = document.createElement("button");
    pin.type = "button";
    pin.className = "ps-pin";
    pin.setAttribute("aria-pressed", "false");
    pin.setAttribute("aria-label", "Pin " + b.name);
    pin.title = "Pin to top";
    pin.textContent = "★";

    var main = document.createElement("div");
    main.className = "ps-main";

    var name = document.createElement("a");
    name.className = "ps-name";
    name.href = pdfUrl;
    name.target = "_blank";
    name.rel = "noopener";
    name.textContent = b.name;

    var meta = document.createElement("div");
    meta.className = "ps-meta";
    var date = document.createElement("span");
    date.className = "ps-date";
    date.textContent = formatDate(b.updated);
    meta.appendChild(date);

    if (b.note) {
      var note = document.createElement("span");
      note.className = "ps-note";
      note.textContent = b.note;
      meta.appendChild(note);
    }

    main.appendChild(name);
    main.appendChild(meta);

    var actions = document.createElement("div");
    actions.className = "ps-actions";
    var btn = document.createElement("a");
    btn.className = "ps-btn";
    btn.href = pdfUrl;
    btn.target = "_blank";
    btn.rel = "noopener";
    btn.setAttribute("tabindex", "-1");
    btn.innerHTML =
      "Open<span class=\"ps-sr\"> " + escapeHtml(b.name) + " price sheet</span>";

    actions.appendChild(btn);

    li.appendChild(pin);
    li.appendChild(main);
    li.appendChild(actions);
    return li;
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function init(builders) {
    var root = document.getElementById("ps");
    if (!root) return;

    builders = (builders || []).slice().sort(function (a, b) {
      var ka = sortKey(a.name);
      var kb = sortKey(b.name);
      if (ka < kb) return -1;
      if (ka > kb) return 1;
      return 0;
    });

    var q = document.getElementById("ps-q");
    var bar = document.getElementById("ps-bar");
    var count = document.getElementById("ps-count");
    var empty = document.getElementById("ps-empty");
    var zero = document.getElementById("ps-zero");
    var pinSec = document.getElementById("ps-pinned");
    var pinList = document.getElementById("ps-pinned-list");
    var groupsEl = document.getElementById("ps-groups");
    var az = document.getElementById("ps-az");

    var pins = loadPins();
    var state = { pinned: false, recent: false };
    var now = new Date();
    var DAY = 86400000;

    if (!builders.length) {
      zero.hidden = false;
      empty.hidden = true;
      pinSec.hidden = true;
      groupsEl.innerHTML = "";
      count.textContent = "0 builders";
      bar.querySelectorAll(".ps-chip").forEach(function (c) {
        c.disabled = true;
      });
      q.disabled = true;
      return;
    }

    zero.hidden = true;

    var letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
    az.innerHTML = "";
    letters.forEach(function (L) {
      var has = builders.some(function (b) {
        return letterFor(b.name) === L;
      });
      if (has) {
        var a = document.createElement("a");
        a.href = "#ps-" + L;
        a.dataset.l = L;
        a.textContent = L;
        az.appendChild(a);
      } else {
        var span = document.createElement("span");
        span.setAttribute("aria-hidden", "true");
        span.textContent = L;
        az.appendChild(span);
      }
    });

    var byLetter = {};
    builders.forEach(function (b) {
      var L = letterFor(b.name);
      if (!byLetter[L]) byLetter[L] = [];
      byLetter[L].push(b);
    });

    groupsEl.innerHTML = "";
    Object.keys(byLetter)
      .sort()
      .forEach(function (L) {
        var sec = document.createElement("section");
        sec.className = "ps-group";
        sec.id = "ps-" + (L === "#" ? "misc" : L);
        sec.dataset.l = L;

        var h = document.createElement("h2");
        h.className = "ps-h";
        h.textContent = L === "#" ? "#" : L;

        var ul = document.createElement("ul");
        ul.className = "ps-list";
        byLetter[L].forEach(function (b) {
          ul.appendChild(renderBuilderItem(b));
        });

        sec.appendChild(h);
        sec.appendChild(ul);
        groupsEl.appendChild(sec);
      });

    var items = [].slice.call(groupsEl.querySelectorAll(".ps-item"));
    var groups = [].slice.call(groupsEl.querySelectorAll(".ps-group"));

    function parseU(u) {
      return u && u.length === 8
        ? new Date(+u.slice(0, 4), +u.slice(4, 6) - 1, +u.slice(6, 8))
        : null;
    }

    items.forEach(function (li) {
      var d = parseU(li.dataset.u);
      var el = li.querySelector(".ps-date");
      li._age = d ? (now - d) / DAY : 1e9;
      if (el && li._age <= 45) el.classList.add("is-new");
    });

    function setTop() {
      root.style.setProperty("--ps-top", getComputedStyle(document.documentElement).getPropertyValue("--header-h").trim() || "72px");
      root.style.setProperty("--ps-bar-h", bar.offsetHeight + 8 + "px");
    }

    function renderPins() {
      pinList.innerHTML = "";
      items.forEach(function (li) {
        var on = pins.indexOf(li.dataset.id) > -1;
        li.querySelector(".ps-pin").setAttribute("aria-pressed", on);
        if (on) {
          var c = li.cloneNode(true);
          c.classList.remove("is-first");
          c.removeAttribute("hidden");
          pinList.appendChild(c);
        }
      });
      pinSec.hidden = !pins.length || state.pinned;
    }

    function apply() {
      var t = norm(q.value);
      var n = 0;
      var first = null;
      items.forEach(function (li) {
        var ok =
          (!t || li.dataset.s.indexOf(t) > -1) &&
          (!state.pinned || pins.indexOf(li.dataset.id) > -1) &&
          (!state.recent || li._age <= 60);
        li.hidden = !ok;
        li.classList.remove("is-first");
        if (ok) {
          n++;
          if (
            t &&
            (!first ||
              (li.dataset.s.indexOf(t) === 0 &&
                first.dataset.s.indexOf(t) !== 0))
          ) {
            first = li;
          }
        }
      });
      if (first) first.classList.add("is-first");

      groups.forEach(function (g) {
        var vis = g.querySelector(".ps-item:not([hidden])");
        g.hidden = !vis;
        var a = az.querySelector('a[data-l="' + g.dataset.l + '"]');
        if (a) a.classList.toggle("is-off", !vis);
      });

      var filtered = !!(t || state.pinned || state.recent);
      count.textContent = filtered
        ? n + " of " + items.length + " builders"
        : items.length + " builders";
      empty.hidden = n > 0;
      pinSec.hidden = !pins.length || filtered;
      root._first = first;
    }

    root.addEventListener("click", function (e) {
      var pinBtn = e.target.closest(".ps-pin");
      if (pinBtn) {
        var li = pinBtn.closest(".ps-item");
        var id = li.dataset.id;
        var i = pins.indexOf(id);
        if (i > -1) pins.splice(i, 1);
        else pins.push(id);
        savePins(pins);
        renderPins();
        apply();
        return;
      }
      var chip = e.target.closest(".ps-chip");
      if (chip) {
        var on = chip.getAttribute("aria-pressed") !== "true";
        chip.setAttribute("aria-pressed", on);
        if (chip.dataset.f) state[chip.dataset.f] = on;
        apply();
      }
    });

    q.addEventListener("input", apply);
    q.addEventListener("keydown", function (e) {
      if (e.key === "Enter") {
        var f =
          root._first ||
          root.querySelector("#ps-groups .ps-item:not([hidden])");
        if (f) {
          e.preventDefault();
          window.open(f.querySelector(".ps-name").href, "_blank", "noopener");
        }
      }
      if (e.key === "Escape") {
        q.value = "";
        apply();
      }
    });

    document.addEventListener("keydown", function (e) {
      if (
        e.key === "/" &&
        document.activeElement !== q &&
        !/input|textarea|select/i.test(document.activeElement.tagName)
      ) {
        e.preventDefault();
        q.focus();
        q.select();
      }
    });

    window.addEventListener("resize", setTop);
    function stuck() {
      var top =
        parseFloat(root.style.getPropertyValue("--ps-top")) ||
        parseFloat(getComputedStyle(root).getPropertyValue("--ps-top")) ||
        72;
      bar.classList.toggle(
        "is-stuck",
        bar.getBoundingClientRect().top <= top + 1 &&
          window.scrollY > 0 &&
          root.getBoundingClientRect().top < top
      );
    }
    window.addEventListener("scroll", stuck, { passive: true });

    setTop();
    renderPins();
    apply();
    stuck();
  }

  fetch("data/builders.json", { cache: "no-cache" })
    .then(function (r) {
      if (!r.ok) throw new Error("load failed");
      return r.json();
    })
    .then(init)
    .catch(function () {
      init([]);
    });
})();
