
function doWidgetSearch(searchPhrase, deptId) {
	if (searchPhrase != '') {
		if (deptId > 0) {
			window.location.href = "https://www.cityofdelafieldwi.gov/Search?searchPhrase=" + encodeURIComponent(searchPhrase) + "&departmentId=" + encodeURIComponent(deptId);
		} else {
			window.location.href = 'https://www.cityofdelafieldwi.gov/Search?searchPhrase=' + encodeURIComponent(searchPhrase);
		}
	}
}

function createDebouncedAutoCompleteSource(options) {
	var currentRequest = null;
	var debounceTimer = null;
	var pendingResponse = null;
	var delay = options.delay || 500;
	var timeout = options.timeout || 10000;

	return function (request, response) {
		if (pendingResponse != null) {
			pendingResponse([]);
			pendingResponse = null;
		}

		clearTimeout(debounceTimer);
		if (currentRequest != null) {
			currentRequest.abort();
			currentRequest = null;
		}

		if (!options.shouldSearch(request.term)) {
			response([]);
			return;
		}

		pendingResponse = response;
		debounceTimer = setTimeout(function () {
			currentRequest = $.ajax({
				url: options.url(),
				dataType: "json",
				timeout: timeout,
				data:
				{
					term: request.term,
				},
				success: function (data) {
					if (pendingResponse === response) {
						pendingResponse = null;
						response(data);
						if (options.success) {
							options.success(data);
						}
					}
				},
				error: function (xmlhttprequest, textstatus, message) {
					if (textstatus === "abort") {
						return;
					}

					if (options.error) {
						options.error(xmlhttprequest, textstatus, message);
					}

					if (pendingResponse === response) {
						pendingResponse = null;
						response([]);
					}
				},
				complete: function (xmlhttprequest) {
					if (currentRequest === xmlhttprequest) {
						currentRequest = null;
					}
				}
			});
		}, delay);
	};
}

//Prevents issue that happens on ios devices where search icon will not load after searching and hitting back button
if (/iPhone|iPad|iPod/i.test(navigator.userAgent)) {
    $( document ).ready(function() {
		$('.imageHover').removeClass('imageHover');
	});
}
