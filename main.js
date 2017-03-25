(function () {
  // Generate excel sheet
  var noOfRows = 5;
  var noOfColumns = 10;
  var cellsId = 1;
  function genExcel(rows, columns) {
    var e1, e2, header, headerRow, row;
    e1 = document.getElementById('table');
    e2 = document.getElementById('tbody');
    header = e1.createTHead();
    headerRow = header.insertRow(0);

    for (var i = 0; i < rows; i++) {
      row = e2.insertRow(i);
      row.className = 'row';
      for (var x = 0; x <= columns; x++) {
        // insert table head
        i === 0 ? createTHead(headerRow, x) : '';

        // insert table cells
        createTBody(row, i, x);
      }
    }
  }

  function createTHead(headerRow, cols) {
    var cell;
    if (cols === 0) {
      headerRow.innerHTML = '<td class="column" height="35"></td>';
    } else {
      cell = headerRow.insertCell(cols);
      cell.className = "column";
      cell.height = 35;
      cell.id = 'col_'+cols;
      cell.innerHTML = cols+'<span class="btn add">+</span>' + 
        '<span class="btn delete">-</span>';
    }
    headerRow.addEventListener('click', addOrDeleteColumn);
  }

  function createTBody(row, i, cols) {
    var th, span, cell;
    if (cols === 0) {
      row.innerHTML = '<th class="column" id="'+(i + 1)+'">' + (i + 1) +
        '<span class="btn add">+</span>' + 
        '<span class="btn delete">-</span>' + 
        '</th>';
      row.addEventListener('click', addOrDeleteRow);
    } else {
      cell = row.insertCell(cols);
      cell.className = "column";
      cell.id = 'cell_'+(cellsId++);
      cell.height = 35;
      cell.innerHTML = '<input value="" readonly tabIndex="1">';
      row.addEventListener('click', editCells);
    }
  }

  genExcel(noOfRows, noOfColumns);

  // Add new row or delete existing row
  function addOrDeleteRow(e) {
    var tbody, insertPos, deletePos, newRow;
    tbody = document.getElementById('tbody');
    if (e.target.className.indexOf('add') >= 0) {
      insertPos = parseInt(e.target.parentElement.id);
      newRow = tbody.insertRow(insertPos);
      newRow.className = 'row';
      // add new clomuns to new row
      for (var x = 0; x <= noOfColumns; x++) {
        createTBody(newRow, insertPos, x);
      }
      updateExistingRows(insertPos, 'add');
      noOfRows++;
    } else if (e.target.className.indexOf('delete') >= 0) {
      deletePos = parseInt(e.target.parentElement.id) - 1;
      tbody.deleteRow(deletePos);
      updateExistingRows(deletePos, 'delete');
      noOfRows--;
    }
  }

  // update existing row
  function updateExistingRows(pos, action) {
    var allRows, nextRow, newRowText, loopLength;
    allRows = document.querySelectorAll('tbody tr');
    for (var i = pos; i < allRows.length; i++) {
      if (action === 'add') {
        nextRow = allRows[i + 1];
        if (nextRow) {
          newRowText = nextRow.querySelector('th').childNodes[0];
          nextRow.querySelector('th').id = parseInt(newRowText.data) + 1;
          newRowText.nodeValue = parseInt(newRowText.data) + 1;
        }
      } else {
        nextRow = allRows[i];
        newRowText = nextRow.querySelector('th').childNodes[0];
        nextRow.querySelector('th').id = parseInt(newRowText.data) - 1;
        newRowText.nodeValue = parseInt(newRowText.data) - 1;
      }
    }
  }

  // Add or delete new column
  function addOrDeleteColumn(e) {
    var table, insertPos, deletePos, newColumn, nextRow, newRowText;
    table = document.getElementById('table');
    if (e.target.className.indexOf('add') >= 0) {
      insertPos = parseInt(e.target.parentElement.id.split('col_')[1]) + 1;
      createTHead(table.querySelector('thead tr'), insertPos);
      for (var i = 0; i < noOfRows; i++) {
        nextRow = table.querySelectorAll('tbody tr')[i];
        newColumn = nextRow.insertCell(insertPos);
        newColumn.className = 'column';
        newColumn.height = '35';
        newColumn.id = 'cell_'+cellsId++;
        newColumn.innerHTML = '<input value="" readonly tabIndex="1" />';
      }
      updateExistingColumns(insertPos, 'add');
      noOfColumns++;
    } else if (e.target.className.indexOf('delete') >= 0) {
      deletePos = parseInt(e.target.parentElement.id.split('col_')[1]);
      // delete td elements
      for (var i = 0; i < noOfRows; i++) {
        nextRow = table.querySelectorAll('tbody tr')[i];
        newColumn = nextRow.deleteCell(deletePos);
      }
      // delete th from header
      table.querySelector('thead tr').removeChild(table.querySelectorAll('thead td')[deletePos - 1]);
      updateExistingColumns(deletePos, 'delete');
      noOfColumns--;
    }
  }

  // update existing column
  function updateExistingColumns(pos, action) {
    var allColumns, nextColumn, newColumnText, loopLength, newPos;
    allColumns = document.querySelectorAll('thead tr td');
    if (action === 'add') {
      loopLength = allColumns.length;
      newPos = pos + 1;
    } else {
      loopLength = allColumns.length + 1;
      newPos = pos;
    }
    for (var i = newPos; i < loopLength; i++) {
      if (action === 'add') {
        nextColumn = allColumns[i];
        newColumnText = nextColumn.childNodes[0];
        nextColumn.id = 'col_'+(parseInt(newColumnText.data) + 1);
        newColumnText.nodeValue = parseInt(newColumnText.data) + 1;
      } else {
        nextColumn = allColumns[i - 1];
        newColumnText = nextColumn.childNodes[0];
        nextColumn.id = 'col_'+(parseInt(newColumnText.data) - 1);
        newColumnText.nodeValue = parseInt(newColumnText.data) - 1;
      }
    }
  }

  // Add cells editable
  var previousCell;
  function editCells(e) {
    if (e.target.nodeName === 'INPUT' || e.target.nodeName === 'TD') {
      var input, elem;
      // make previous input as readonly
      if (previousCell) {
        previousCell.setAttribute('readonly', true);
        previousCell.parentElement.classList.remove('editable');
      }
      // Fetch target cell
      if (e.target.nodeName === 'INPUT') {
        elem = document.getElementById(e.target.parentNode.id);
      } else if (e.target.nodeName === 'TD') {
        elem = document.getElementById(e.target.id);
      }
      // Make cell editable
      elem.classList.add('editable');
      elem.children[0].removeAttribute('readonly');
      elem.children[0].focus() // add auto focus on click
      previousCell = elem.children[0]; // store reference of previous cell input
    }
  }

  // Make tab work
  document.addEventListener('keyup', function(e) {
    if (e.keyCode === 9) {
      editCells(e);
    }
  });

})();
