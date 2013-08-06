
Converter = function() {}

Converter.prototype.toJSON = function(workbook) {
  var result = {};
  workbook.SheetNames.forEach(function(sheetName) {
    var roa = XLS.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
    if (roa.length > 0){
      result[sheetName] = roa;
    }
  });
  return result;
}

Converter.prototype.process = function(workbook) {
  json = this.toJSON(workbook);

  converter.errors = [];
  converter.output = null;
  converter.choiceLists = {};

  form = { elements: [] };

  this.createChoices(json.choices);

  var self = this;
  mainSheetName = workbook.SheetNames[0];
  _.each(json[mainSheetName], function(row) {
    element = self.createElement(row, json);
    if (element) {
      form.elements.push(element);
    }
  });

  output = { form: form };

  formName = '';

  if (json.settings && json.settings.length > 0) {
    formName = json.settings[0].form_title;
  } else if (converter.fileName) {
    formName = converter.fileName.replace('.xls', '').replace('.xlsx', '');
  }

  form.name = form.description = formName || 'XLSForm';

  converter.output = output;

  document.querySelector('#form').innerHTML = JSON.stringify(output, 2, 2);

  converter.updateErrors();
}

Converter.prototype.createChoices = function(choices) {
  converter.choiceLists = {}
  _.each(choices, function(choice) {
    listNameAttribute = choice.list_name ? 'list_name' : 'list name'
    choiceListName = choice[listNameAttribute];
    converter.choiceLists[choiceListName] = converter.choiceLists[choiceListName] || []
    if (choice.label) {
      converter.choiceLists[choiceListName].push(
        {
          label: choice.label.toString(),
          value: choice.name.toString()
        }
      );
    }
  });
}

Converter.prototype.createElement = function(row, workbook) {
  field = null;

  if (!row.type)
    return null;

  parts = row.type.split(' ');
  type = parts[0];
  choiceList = null;
  allowOther = null;

  if (parts.length > 1) {
    choiceList = parts[1];
  }
  if (parts.length > 2) {
    allowOther = parts[2] === 'or_other';
  }

  switch (type) {
    case 'text':
      field = {
        type: 'TextField',
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || ''
      };

      break;

    case 'integer':
      field = {
        type: 'TextField',
        numeric: true,
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || ''
      };

      break;

    case 'decimal':
      field = {
        type: 'TextField',
        numeric: true,
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || ''
      };

      break;

    case 'select_one':
      field = {
        type: 'ChoiceField',
        allow_other: allowOther === true,
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || '',
        choices: converter.choiceLists[choiceList]
      };


      break;

    case 'acknowledge':
      field = {
        type: 'ChoiceField',
        allow_other: allowOther === true,
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || '',
        choices: [ { label: 'Yes' }, { label: 'No' } ]
      };


      break;

    case 'select_multiple':
      field = {
        type: 'ChoiceField',
        multiple: true,
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || '',
        choices: converter.choiceLists[choiceList]
      };

      break;


    case 'note':
      field = {
        type: 'Label',
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || ''
      };

      break;

    case 'image', 'photo':
      field = {
        type: 'PhotoField',
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || ''
      };
      break;

    case 'date':
      field = {
        type: 'DateTimeField',
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || ''
      };

      break;

    case 'datetime':
      field = {
        type: 'DateTimeField',
        data_name: row.name,
        key: row.name,
        label: row.label,
        description: row.hint || ''
      };

      break;

    default:
      converter.errors.push(row.type + ' field type is not supported.');
      console.log(row.type + ' is not supported.');
      console.log(row);
      break
  }

  if (field) {
    field.required = row.required === 'yes' ? true : false;
    field.hidden = false;
    field.disabled = false;
  }

  return field;
}

Converter.prototype.initialize = function() {
  var drop = document.getElementById('drop');
  if (drop.addEventListener) {
    drop.addEventListener('dragenter', this.handleDragover, false);
    drop.addEventListener('dragover', this.handleDragover, false);
    drop.addEventListener('drop', this.handleDrop, false);
  }
  var button = document.getElementById('upload');
  button.addEventListener('click', this.handleUpload, false);
  document.getElementById('token').value = Cookies.get('FulcrumToken') || '';
}

Converter.prototype.handleUpload = function(e) {
  if (!converter.output) {
    alert('You must first drag an XLS file.');
    return;
  }

  if (document.getElementById('token').value.length === 0) {
    alert('You must enter your API token.');
    return;
  }

  Cookies.set('FulcrumToken', document.getElementById('token').value);

  superagent
    .post('https://api.fulcrumapp.com/api/v2/forms')
    .send(converter.output)
    .set('X-ApiToken', document.getElementById('token').value)
    .set('Accept', 'application/json')
    .end(function(error, res){
      if (error || res.status !== 201) {
        alert('There was an error creating the form in your Fulcrum account.');
        converter.errors.push(res.text);
        converter.updateErrors();
        console.log(error);
        console.log(res);
      } else {
        alert('Successfully created form in your Fulcrum account.');
      }
    });
}

Converter.prototype.updateErrors = function() {
  errorContainer = document.querySelector('#errors');

  if (converter.errors.length > 0) {
    errorContainer.style.display = '';
    errorContainer.innerHTML = converter.errors.join("<br />");
  } else {
    errorContainer.style.display = 'none';
  }
}

Converter.prototype.handleDrop = function(e) {
  e.stopPropagation();
  e.preventDefault();
  var files = e.dataTransfer.files;
  var i,f;
  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader();
    var name = f.name;
    converter.fileName = name;
    reader.onload = function(e) {
      var data = e.target.result;
      var cfb = XLS.CFB.read(data, {type: 'binary'});
      var wb = XLS.parse_xlscfb(cfb);
      converter.process(wb);
    };
    reader.readAsBinaryString(f);
  }
}

Converter.prototype.handleDragover = function(e) {
  e.stopPropagation();
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
}

window.onload = function() {
  converter = new Converter();
  converter.initialize();
}
