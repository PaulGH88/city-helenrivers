/* The City's own scripts ask their server for the menus, the mega-menu content, feature toggles
   and a form token as the page runs. This copy has no City server behind it, so those requests
   are answered from files captured from the live homepage at build time (tools/capture-endpoints.mjs).
   Nothing is sent anywhere. Loaded first, before any City script. */
(function () {
	'use strict';
	var MAP = [
		[/\/Toggle$/i, 'cp/toggle.json'],
		[/\/antiforgery$/i, 'cp/antiforgery.json'],
		[/\/Pages\/MenuMain\/HiddenMainSubMenus$/i, 'cp/submenus.html'],
		[/\/Content\/Load$/i, 'cp/content-load.html']
	];
	function local(url) {
		try {
			var u = new URL(url, location.href);
			if (u.origin !== location.origin) return null;
			for (var i = 0; i < MAP.length; i++) if (MAP[i][0].test(u.pathname)) return MAP[i][1];
		} catch (e) {}
		return null;
	}
	var open = XMLHttpRequest.prototype.open;
	XMLHttpRequest.prototype.open = function (method, url) {
		var to = local(url);
		var args = Array.prototype.slice.call(arguments);
		if (to) { args[0] = 'GET'; args[1] = to; }
		return open.apply(this, args);
	};
	// Safety net: a City link that their scripts point at a site-root path (search, menus built
	// at run time) goes to the real City site, never to a page this copy does not have.
	var OWN = /^\/(index\.html)?$|^\/(helen|cp)\//;
	document.addEventListener('click', function (e) {
		var a = e.target && e.target.closest ? e.target.closest('a[href]') : null;
		if (!a) return;
		var u; try { u = new URL(a.getAttribute('href'), location.href); } catch (err) { return; }
		if (u.origin === location.origin && !OWN.test(u.pathname) && a.getAttribute('href').charAt(0) !== '#') {
			a.href = 'https://www.cityofdelafieldwi.gov' + u.pathname + u.search + u.hash;
		}
	}, true);
	// Anything else of theirs that fails here would pop an alert box written for the City's
	// editors; on this copy it goes to the console instead.
	window.alert = function (m) { if (window.console) console.warn('[city script message]', m); };
	window.confirm = function (m) { if (window.console) console.warn('[city script confirm]', m); return false; };
})();
