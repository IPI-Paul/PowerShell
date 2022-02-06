const selected = {
  'Id': [],
  'Contact_Title': [],
  'First_Name': [],
  'Last_Name': [],
  'Phones': [],
  'Emails': []
};

function clearFilter() {
  $('table tr').each(function() {
    this.style.display = '';
  });
}

function filterRows() {
  $('table tr').each(function($idx) {
    var filter = false;
    $(this).find("td").each(function () {
      if ($(this)[0].style.backgroundColor == "rgb(253, 233, 217)")  {
        filter = true;
      }
    });
    if (!filter) {
      if ($idx > 0) {
        this.style.display = 'none';
      }
    }
  });
}

function clearHighlight() {
  $('table tr td').each(function() {
    if (this.style.backgroundColor == "rgb(253, 233, 217)")  {
      this.style.backgroundColor = "rgb(255, 255, 255)";
    }
  });
}

function getHTML() {
  var style = '<style>';
  $.each(document.styleSheets, function(sheetIndex, sheet) {
    if (sheet.media == 'screen'){
      $.each(sheet.cssRules || sheet.rules, function(ruleIndex, rule) {
        style = style + '\n' + rule.cssText;
      });
    }
  });
  style = style + '\n</style>\n';
  return style + document.body.outerHTML.replace('<table> </table>', '');
}

$(document).ready(function() {
  var br = (document.location.href == 'about:blank') ? '\n' : '<br />';
  $('body').on('click', 'td', function () {
      if ($(this)[0].className != 'func') {
          ColName = $('table tr th')[$(this).parent().children('td').index(this)].innerText.trim().replace(/ /g, '_');
          var tmp = {0: []};
          if ($(this)[0].style.backgroundColor == "rgb(253, 233, 217)") {
              $(this).css('background-color', 'rgb(255, 255, 255)');
              for ($i = 0; $i < selected[ColName].length; $i++) {
                if(selected[ColName][$i] != this.innerText.trim().replace(/\r\n/g, ', ')) {
                  tmp[0].push(selected[ColName][$i]);
                }
              }
              selected[ColName] = [];
              for ($i = 0; $i < tmp[0].length; $i++) {
                selected[ColName].push(tmp[0][$i]);
              }
          } 
          else {
            $(this).css('background-color', 'rgb(253, 233, 217)');
            selected[ColName].push(this.innerText.trim().replace(/\r\n/g, ', '));
          }
          document.title = JSON.stringify(selected);
          window.location.href = '#';
      }
  });
  $('table tr').each(function () {
    $(this).find("th").each(function ($idx) {
      this.innerText = this.innerText.toString().replace(/_/g, ' ');
    });
    $(this).find("td").each(function ($idx) {
      this.innerText = this.innerText.trim().replace(", ", "\n");
        if ($idx == 4) {
            var innr = Array();
            $(this.innerText.trim().split('\n')).each(function () {
              var tel = document.createElement('a');
              tel.href = 'tel:' + this;
              tel.innerText = this;
              innr[innr.length] = tel.outerHTML;
            });
            this.innerHTML = innr.join(br).replace('<br></a>', '</a><br>');
        }
        if ($idx == 5) {
            var innr = Array();
            var sibs = $(this).parent().children('td');
            var fName = sibs[1].innerText;
            var lName = sibs[2].innerText;
            $(this.innerText.trim().split('\n')).each(function () {
              var eml = document.createElement('a');
              eml.href = 'mailto:?to=' + fName + '%20' + lName + '<' + this + '>';
              eml.innerText = this;
              innr[innr.length] = eml.outerHTML;
            });
            this.innerHTML = innr.join(br).replace('<br></a>', '</a><br>');
        }
    });
  });
});