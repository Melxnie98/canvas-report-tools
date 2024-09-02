// ==UserScript==
// @name         Canvas Quiz to excel
// @namespace    https://github.com/sukotsuchido/CanvasUserScripts
// @version      0.1
// @description  Allows the user to print quizzes from the preview page.
// @author        Wen Hol customized to allow for quizbanks
// @include      https://*/courses/*/question_banks/*
// @require     https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js
// @require     https://flexiblelearning.auckland.ac.nz/javascript/filesaver.js
// @require     https://flexiblelearning.auckland.ac.nz/javascript/xlsx.full.min.js
// ==/UserScript==

//------------------------------USER SCRIPT define metadata for the userscript manager and provide essential information about how and when the script should run

// Self-executing anonymous function to avoid polluting the global scope
(function() {
    // Ensure the document is fully loaded before running the script
    $(document).ready(function() {
        // Select the right-side panel of the Canvas page to append the new button
        var parent = document.querySelector('#right-side');

        // Create a new button element
        var el = document.createElement('button');
        el.classList.add('Button', 'element_toggler', 'button-sidebar-wide'); // Add necessary classes for styling
        el.type = 'button'; // Set button type
        el.id = 'printQuizButton'; // Set unique ID for the button

        // Create an icon element and add it to the button
        var icon = document.createElement('i');
        icon.classList.add('icon-document'); // Add class for icon styling
        el.appendChild(icon);

        // Create text node for button label and add it
        var txt = document.createTextNode(' Download to excel');
        el.appendChild(txt);

        // Attach click event listener to button; triggers function to handle quiz export
        el.addEventListener('click', allMatchQuestions);

        // Append the newly created button to the right-side panel
        parent.appendChild(el);
    });

    // Create a temporary table element to store quiz data
    var $tmpTable = $('<table id="tmpTable" />');
    var CRLF = '<br> \r\n'; // Line break for formatting

    // Extract the quiz title from the page
    var quizTitle = jQuery('.quiz-header').find('.displaying').text();
    
    // Get the current date to use in the filename
    var today = new Date();
    var dd = today.getDate();
    var mm = today.getMonth() + 1; // Months are zero-based
    var yyyy = today.getFullYear();
    var debug = 0; // Debug flag for logging

    // Format day and month with leading zero if necessary
    if (dd < 10) {
        dd = '0' + dd;
    }
    if (mm < 10) {
        mm = '0' + mm;
    }

    // Format date for filename, including a timestamp
    today = (yyyy - 2000) + '-' + mm + '-' + dd + '-' + Math.floor(Date.now() / 1000);

    // Function to process and format all matching questions
    function allMatchQuestions() {
        // Remove the "brief" class to ensure full question view
        jQuery("#questions").removeClass("brief");
        
        // Hide unnecessary elements like quiz header paragraphs
        jQuery('.quiz-header p').hide();

        // Select all matching questions on the page
        var allMatchQuestions = document.querySelectorAll("div.matching_question");
        for (var z = 0; z < allMatchQuestions.length; z++) {
            var options = allMatchQuestions[z].querySelector("select").options;
            var list = document.createElement('div');
            var matchText = document.createElement('div');
            matchText.style.verticalAlign = 'middle';
            matchText.innerHTML = '<strong>Match Choices:</strong>';
            
            // Create a list of matching choices for each question
            for (var t = 0; t < options.length; t++) {
                if (options[t].textContent !== "[ Choose ]") {
                    var temp = document.createElement('div');
                    temp.innerHTML = options[t].text;
                    temp.style.display = 'inline-block';
                    temp.style.padding = '20px';
                    temp.style.maxWidth = '25%';
                    temp.style.verticalAlign = 'Top';
                    list.appendChild(temp);
                }
                list.style.width = 'inherit';
                list.style.border = "thin dotted black";
                list.style.padding = "0px 0px 0px 10px";

                // Append the list of choices to the question's answer section
                var optionsList = allMatchQuestions[z].querySelector(".answers");
                optionsList.appendChild(matchText);
                matchText.appendChild(list);

                // Hide select elements to avoid duplication in the export
                var hideOptions = allMatchQuestions[z].querySelectorAll("select");
                console.log(hideOptions);
                for (var q = 0; q < hideOptions.length; q++) {
                    hideOptions[q].style.visibility = "hidden";
                }
            }
        }
        // Process multi-select questions and format the quiz for printing/exporting
        multiSelectQuestions();
        printQuizStyle();
    }

    // Function to handle formatting of multi-select questions
    function multiSelectQuestions() {
        var allMultiSelectQuestions = document.querySelectorAll("div.multiple_dropdowns_question select");
        for (var q = 0; q < allMultiSelectQuestions.length; q++) {
            var len = allMultiSelectQuestions[q].options.length;
            allMultiSelectQuestions[q].setAttribute('size', len); // Set size to display all options
            allMultiSelectQuestions[q].style.width = 'fit-content'; // Adjust width to fit content
            allMultiSelectQuestions[q].style.maxWidth = ''; // Remove any max-width constraints
        }
    }

    // Function to apply styles to prepare the quiz for export and printing
    function printQuizStyle() {
        var scale = document.querySelector("div.ic-Layout-contentMain");
        scale.style.zoom = "74%"; // Scale down the content for better fit

        // Prevent page breaks within question blocks
        var questionBlocks = document.querySelectorAll("div.question_holder");
        for (var i = 0; i < questionBlocks.length; i++) {
            questionBlocks[i].style.pageBreakInside = "avoid";
        }

        // Style answer choices for better alignment and appearance
        var answerChoices = document.querySelectorAll("div.answer");
        for (var j = 0; j < answerChoices.length; j++) {
            answerChoices[j].style.verticalAlign = "Top";
            answerChoices[j].style.borderTop = "none";
        }

        // Hide form actions, editor tools, and other elements not needed for the export
        var formActions = document.querySelectorAll("div.alert, div.ic-RichContentEditor, div.rce_links");
        for (var h = 0; h < formActions.length; h++) {
            formActions[h].style.visibility = "hidden";
        }
        var essayShrink = document.querySelectorAll("div.mce-tinymce");
        for (var m = 0; m < essayShrink.length; m++) {
            essayShrink[m].style.height = "200px";
        }
        var bottomLinks = document.querySelectorAll(".bottom_links");
        for (var k = 0; k < bottomLinks.length; k++) {
            bottomLinks[k].style.visibility = "hidden";
        }
        var arrowInfo = document.querySelectorAll(".answer_arrow");
        for (var l = 0; l < arrowInfo.length; l++) {
            arrowInfo[l].style.visibility = "hidden";
        }
        var labelDetails = document.querySelectorAll("label[for='show_question_details']");
        for (var l = 0; l < labelDetails.length; l++) {
            labelDetails[l].style.visibility = "hidden";
        }
        jQuery('#show_question_details').hide();

        // Add all questions and their options to the temporary table
        jQuery('.question').each(function() {
            if (jQuery(this).find('.question_text').text().trim() != "") {
                jQuery(this).find('.header').each(function() {
                    $tmpTable.append("<tr><td>" + jQuery(this).text() + "</td></tr>");
                });
                let tmpQtext = jQuery(this).find('.question_text').html();
                let tmpOptions = '';
                jQuery(this).find('.answer_text').each(function() {
                    tmpOptions += '&nbsp;&nbsp;&nbsp;&nbsp;' + jQuery(this).text() + CRLF;
                });
                $tmpTable.append("<tr><td>" + tmpQtext + CRLF + tmpOptions + "</td></tr>");
                if (debug) console.log(tmpQtext, tmpOptions);
            }
        });

        // Call function to export the table data to an Excel file
        ExportToExcel();
    }

    // Function to export the quiz data from the temporary table to an Excel file
    function ExportToExcel() {
        jQuery('body').append($tmpTable); // Append the temporary table to the body
        var elt = document.getElementById('tmpTable'); // Get the temporary table element
        var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" }); // Convert table to Excel workbook

        // Set properties for the workbook
        wb.Props = {
            Title: quizTitle, // Title of the workbook
            Subject: "", // Subject of the workbook
            Author: "", // Author of the workbook
            CreatedDate: new Date() // Creation date of the workbook
        };

        // Write the workbook to a binary string
        let wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
        let blob = new Blob([s2ab(wbout)], { 'type': 'application/octet-stream' }); // Create a Blob object for the binary data

        // Define the filename for the downloaded Excel file
        let savename = 'quizbank' + quizTitle + '-' + today + '.xlsx';
        saveAs(blob, savename); // Trigger download of the file
        jQuery('#tmpTable').remove(); // Clean up temporary table from the DOM
    }

    // Utility function to convert a string to an ArrayBuffer
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length); // Create an ArrayBuffer with the length of the string
        var view = new Uint8Array(buf); // Create a view for the ArrayBuffer
        for (var i = 0; i < s.length; i++) {
            view[i] = s.charCodeAt(i) & 0xFF; // Convert string characters to byte values
        }
        return buf; // Return the ArrayBuffer
    }
})();
