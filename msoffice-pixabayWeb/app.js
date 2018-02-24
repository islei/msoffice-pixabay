(function () {

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {

        "use strict";

        var CONFIG = {
            API_KEY: "YOUR_API_KEY",
            PIXABAY_URL: "https://pixabay.com/api/",
            RESULTS_PER_PAGE: 20,
        }

        // initialize search box
        var SearchBoxElement = document.querySelector(".ms-SearchBox");
        var searchBox = new fabric['SearchBox'](SearchBoxElement);

        // initialize spinner
        var SpinnerElement = document.querySelector('.ms-Spinner');
        var spinner = new fabric['Spinner'](SpinnerElement);

        // on search
        $('#pxb-searchBox-field').keyup(function (event) {
            if (event.keyCode == 13) {
                var searchVal = $(this).val();
                if (!searchVal) return;
                search(searchVal);
            }
        });

         /**
         * Search for images on pixabay using the keyword(s)
         */
        function search(searchVal, options) {
            // get options
            var page = options && options.page ? options.page : 1;
            // send
            $.ajax({
                url: CONFIG.PIXABAY_URL,
                data: {
                    key: CONFIG.API_KEY,
                    per_page: CONFIG.RESULTS_PER_PAGE,
                    page: page,
                    q: searchVal,
                },
                beforeSend: function () {
                    $("#pxb-result-container > div").empty();
                    $("#pxb-spinner-container").css("display", "block");
                },
                success: function (response) {
                    var hits = response.hits;
                    var totalHits = response.totalHits;
                    // populate result container with preview images
                    $.each(hits, function (i, hit) {
                        var previewImg = $('<img src="' + hit.previewURL + '" />');
                        // on preview image click
                        previewImg.click(function () {
                            insertImage(hit.webformatURL);
                        })
                        var selector = (i & 1) ? "first" : "last";
                        selector = "#pxb-result-container > div:" + selector + "-child";
                        $(selector).append(previewImg);
                    })
                    // set pagination
                    if (page != 1) return;
                    $("#pxb-pagination").paging(totalHits, {
                        format: '< ncnnn >',
                        perpage: CONFIG.RESULTS_PER_PAGE,
                        onSelect: function (page) {
                            if (page == 1) return;
                            search(searchVal, { page: page })
                        },
                        onFormat: formatPaging,
                    });
                },
                complete: function () {
                    // hide spinner
                    $("#pxb-spinner-container").css("display", "none");
                }
            })
        }

        /**
         * Paging button formatter
         */
        function formatPaging(type) {
            switch (type) {
                case 'block': // n and c
                    if (this.value != this.page)
                        return '<a>' + this.value + '</a>';
                    return '<a class="active">' + this.value + '</a>';
                case 'next': // >
                    return '<a>&gt;</a>';
                case 'prev': // <
                    return '<a>&lt;</a>';
            }
        }

        /**
         * Inserts an image at the cursor position
         */
        function insertImage(imageUrl) {
            try {
                toDataURL(imageUrl, function (dataUrl) {
                    var base64result = dataUrl.split(',')[1];
                    Office.context.document.setSelectedDataAsync(base64result, {
                        coercionType: Office.CoercionType.Image,
                        imageLeft: 50,
                        imageTop: 50,
                        imageWidth: 400,
                    })
                })
            }
            catch (exception) {
                OfficeHelpers.Utilities.log(exception);
            }
        }

        /**
         * convert image to base64 data uri
         */
        function toDataURL(url, callback) {
            var xhr = new XMLHttpRequest();
            xhr.onload = function () {
                var reader = new FileReader();
                reader.onloadend = function () {
                    callback(reader.result);
                }
                reader.readAsDataURL(xhr.response);
            };
            xhr.open('GET', url);
            xhr.responseType = 'blob';
            xhr.send();
        }

    };

})();
