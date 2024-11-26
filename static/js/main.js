$(function() { 
    var server_url = "https://api.bytesview.com";
    var client = ZAFClient.init();
    client.metadata().then(function(metadata) {
      app_status = metadata.settings.app_status
      token = metadata.settings.token
      api_key = metadata.settings.api_key
  
      if (app_status == null){
        showAPIkeypage(client, server_url);
      }
      else if (app_status == "showOauthpage"){
        showOauthpage(client, server_url);
      }
      else if (app_status == "showFieldModel"){
        showFieldModel(client, token, server_url);  
      }    
      else{
        showfinalWindow(client, token, server_url);
      }
    });
  
  });
  
  
  function showAPIkeypage(client, server_url) {
    var source = $("#api_key-hdbs").html();
    var template = Handlebars.compile(source);
    var html = template();
    $("#content").html(html);
  
    $("#check-api_key").click(function (event) {
      event.preventDefault();
      var api_key = document.getElementById("api_key").value;
      var dict = {};
      dict["apikey"] = api_key;
  
      var settings = {
        url: 'https://api.bytesview.com/verify/apikey',
        type: 'POST',
        cors: false,
        contentType: 'application/json',
        data: JSON.stringify(dict)
      };
      client.request(settings).then(
        function (data) {
          if (data['0']==1){
            client.invoke('notify', 'API key Verification completed');
            installationAPIkeyflag(client, api_key, server_url);
          }
          else{
            client.invoke('notify', 'API key Verification needed!', 'error');
          }
        },
        function (response) {
          var msg = 'Error ' + response.status;
          client.invoke('notify', msg, 'error');
        }
      );
    });
  }
  
  
  function installationAPIkeyflag(client, api_key, server_url){
    var settings = {
    url: '/api/v2/apps/installations.json',
    type: 'GET',
    headers:{"Content-Type": "application/json"}
  };
  client.request(settings).then(
    function (data) {
      for (i = 0; i < data["installations"].length; i++) {
        if (data["installations"][i]["settings"]["app_id"] == 123456789){
          var settings = {
            url: '/api/v2/apps/installations/'+data["installations"][i]["id"]+'.json',
            type: 'PUT',
            headers:{"Content-Type": "application/json"},
            data: JSON.stringify({"settings": {"app_status": "showOauthpage", "api_key": api_key}})
          };
          client.request(settings).then(
            function (response) {
              showOauthpage(client, server_url);
            },
            function (response) {
              var msg = 'Error ' + response.status;
              client.invoke('notify', msg, 'error');
            }
            )          
        }
    }
    });
  }
  
  
  function showOauthpage(client, server_url) {
  
    client.context().then(function(context) {
      subdomain = context.account.subdomain;
      var source = $("#oauth_page-hdbs").html();
      var template = Handlebars.compile(source);
      var html = template({user_name: subdomain, client_id:"zdg-bytesview_zendesk_app", redirect_uri: server_url + '/zendesk/handle_user_decision', state: subdomain});
      $("#content").html(html);
    });
  }
  
  
  function showFieldModel(client, token, server_url) {
    client.context().then(function(context) {
    subdomain = context.account.subdomain;
    dict_token = {};
    dict_token["access_token"] = token;
    dict_token["subdomain"] = subdomain;
    var settings = {
        url: server_url + '/zendesk/get_selected_field_option',
        type: 'POST',
        contentType: 'application/json',
        data: JSON.stringify(dict_token)
    };
    client.request(settings).then(
        function (data) {
        showFieldModel1(client, token, data, server_url)
        },
        function (response) {
            var msg = 'Error ' + response.status;
            client.invoke('notify', msg, 'error');
        }
    );
    });
  }
  
  
  function showFieldModel1(client, token, options, server_url) {
    var selectedField = options["selected_field"]
    var modelOptions = options["model_option"]
    var fieldOptions = options["custom_ticket_field"]
    //var selectedField = {"0":{"model":"sentiment","field": "sentiment(1997837)"},"1": {"model":"keywords", "field": "sentiment(199678)"}}
    //var modelOptions = ["sentiment", "keywords", "emotion"]
    //var fieldOptions = ["sentiment(1997837)", "sentiment(199678)", "emotion(199678)"]
    
  
    var modelOptions_string = "";
    var fieldOptions_string = "";
  
    modelOptions_string = modelOptions_string + '<option disabled="disabled" value="disabled" selected>Choose a model...</option>';
    for (var i = 0; i < modelOptions.length; i++) {
      modelOptions_string += "<option>" + modelOptions[i] + "</option>";
    }
  
    for (var i = 0; i < fieldOptions.length; i++) {
      fieldOptions_string += "<option>" + fieldOptions[i] + "</option>";
    }
  
    var source = $("#field_model-hdbs").html();
    var template = Handlebars.compile(source);
    var html = template();
    $("#content").html(html);
  
    if (Object.keys(selectedField).length >= 1){
      for (var i = 0; i < Object.keys(selectedField).length; i++) {
        var field_id = Object.keys(selectedField)[i]
        var row_value = Object.keys(selectedField)
        //selectedField[Object.keys(selectedField)[row_value[i]]]
        var a = selectedField[Object.keys(selectedField)[i]]
        document.getElementById("rowField").appendChild(htmlToElement('<div id=row_id'+ field_id +' class="combination-row combination-row-notError" ><div class="row"><div class="col-xs-6 first-column"><span class="combination-label">Use</span><select disabled="disabled" class="form-control" id="modelList'+ field_id +'">'+ "<option selected>" + a["model"] + "</option>" +'</select></div><div class="col-xs-6 second-column"><span class="combination-label">in</span><input disabled="disabled" class="form-control" list="datalistOptions'+ field_id +'" id="exampleDataList' + field_id + '" placeholder="Field Name..." value='+ a["field"] +' ><span class="glyphicon glyphicon-trash" id=glyphicon_button_trash'+ field_id +' value=row_id' + field_id+ ' role="button" aria-hidden="true" data-toggle="tooltip" data-placement="right" title="Remove this automation field" style=""></span></div></div></div>'));
      }
    }
    else{
      field_id = 0;    
      document.getElementById("rowField").appendChild(htmlToElement('<div id=row_id'+ field_id +' class="combination-row combination-row-notError" ><div class="row"><div class="col-xs-6 first-column"><span class="combination-label">Use</span><select class="form-control" id="modelList'+ field_id +'">'+ modelOptions_string +'</select></div><div class="col-xs-6 second-column"><span class="combination-label">in</span><input class="form-control" list="datalistOptions'+ field_id +'" id="exampleDataList' + field_id + '" placeholder="Field Name..."><datalist id="datalistOptions'+ field_id +'">'+ fieldOptions_string +'</datalist><span class="glyphicon glyphicon-trash" id=glyphicon_button_trash'+ field_id +' value=row_id' + field_id+ ' role="button" aria-hidden="true" data-toggle="tooltip" data-placement="right" title="Remove this automation field" style=""></span></div></div></div>'));    
      selectedField[field_id] = {}
    }
    //document.getElementById($(this).attr("value")).remove();
    //document.getElementById("rowField").appendChild(htmlToElement('<div id=row_id'+ field_id +' class="combination-row combination-row-notError" ><div class="row"><div class="col-xs-6 first-column"><span class="combination-label">Use</span><select class="form-control" id="modelList'+ field_id +'">'+ modelOptions_string +'</select></div><div class="col-xs-6 second-column"><span class="combination-label">in</span><input class="form-control" list="datalistOptions'+ field_id +'" id="exampleDataList' + field_id + '" placeholder="Field Name..."><datalist id="datalistOptions'+ field_id +'">'+ fieldOptions_string +'</datalist><span class="glyphicon glyphicon-trash" id=glyphicon_button_trash'+ field_id +' value=row_id' + field_id+ ' role="button" aria-hidden="true" data-toggle="tooltip" data-placement="right" title="Remove this automation field" style=""></span></div></div></div>'));
    //document.getElementById("rowField").appendChild(htmlToElement(text));
    //document.getElementById("rowField").innerHTML = text;
  
    //text += '<div class="combination-row combination-row-notError" ><div class="row"><div class="col-xs-6 first-column"><span class="combination-label">Use</span><select class="form-control" id="modelList'+ field_id.toString() +'">'+ "<option selected>" + Object.keys(selectedField)[i] + "</option>" +'</select></div><div class="col-xs-6 second-column"><span class="combination-label">in</span><input class="form-control" list="datalistOptions" id="exampleDataList" placeholder="Field Name..."><datalist id="datalistOptions'+ field_id.toString() +'">'+ a +'</datalist></div></div></div>'
    // document.getElementById("datalistOptions").innerHTML = catOptions;
    // document.getElementById("modelList").innerHTML = modelOptions;
  
    var j = 0;
    for (var i = 0; i < Object.keys(selectedField).length; i++) {
      j = Object.keys(selectedField)[i]
      $("#glyphicon_button_trash"+j).click(function (event) {
        event.preventDefault();
        document.getElementById($(this).attr("value")).remove();
        var delete_row = $(this).attr("value").match(/(\d+)/)['0']
        delete selectedField[delete_row];
      })
    }
    
    
  
    $("#glyphicon_button").click(function (event) {
      var field_id;
      var fields = Object.keys(selectedField);
      if (fields.length != 0){
        fields = fields.map(i=>Number(i))
        field_id = Math.max.apply(Math, fields);
      }
      else{
        field_id = -1
      }
      field_id = field_id+1
      
      document.getElementById("rowField").appendChild(htmlToElement('<div id=row_id'+ field_id +' class="combination-row combination-row-notError" ><div class="row"><div class="col-xs-6 first-column"><span class="combination-label">Use</span><select class="form-control" id="modelList'+ field_id +'">'+ modelOptions_string +'</select></div><div class="col-xs-6 second-column"><span class="combination-label">in</span><input class="form-control" list="datalistOptions'+ field_id +'" id="exampleDataList' + field_id + '" placeholder="Field Name..."><datalist id="datalistOptions'+ field_id +'">'+ fieldOptions_string +'</datalist><span class="glyphicon glyphicon-trash" id=glyphicon_button_trash'+ field_id +' value=row_id' + field_id+ ' role="button" aria-hidden="true" data-toggle="tooltip" data-placement="right" title="Remove this automation field" style=""></span></div></div></div>'));
      selectedField[field_id] = {}
  
      $("#glyphicon_button_trash"+field_id).click(function (event) {
            event.preventDefault();
            document.getElementById($(this).attr("value")).remove();
            var delete_row = $(this).attr("value").match(/(\d+)/)['0']
            delete selectedField[delete_row];
          })
    })
  
  
    $("#check-run").click(function (event) {
      event.preventDefault();
      var validation = 1;    
  
      var models = []
      var fields = []
      var row_value = Object.keys(selectedField)
      for (var i = 0; i < Object.keys(selectedField).length; i++) {
        var a = selectedField[Object.keys(selectedField)[i]]
        
        if (Object.keys(a).length === 0 && a.constructor === Object){
          var e = document.getElementById("modelList"+row_value[i]);
          var selectModel = e.options[e.selectedIndex].value;
          var selectField = document.getElementById("exampleDataList"+row_value[i]).value;
          
          if(( selectModel == "disabled" ) || ( selectField == "" )){
            client.invoke('notify', "The Input Model or Ticket Field empty", 'error');
            validation = 0;
            break;
            }
          else if(models.includes(selectModel)){
            client.invoke('notify', "Duplicate Model selected", 'error');
            validation = 0;
            break;
          }
          else if(fields.includes(selectField)){
            client.invoke('notify', "Duplicate Ticket field selected", 'error');
            validation = 0;
            break;
          } 
          else{
            selectedField[row_value[i]]["model"] = selectModel;
            selectedField[row_value[i]]["field"] = selectField;
          }
          }
          models.push(selectedField[row_value[i]]["model"])
          fields.push(selectedField[row_value[i]]["field"])
        }
      if (Object.keys(selectedField).length == 0){
        client.invoke('notify', "Atleast one Model and Ticket field needed!", 'error');
        validation = 0;
      }
  
        
      if (validation == 1){
        createTicketField(client, token, selectedField, server_url)
      }      
    });
  }
  
  
  function createTicketField(client, token, dict, server_url) {
  
    client.context().then(function(context) {
    subdomain = context.account.subdomain;
    dict_token = {};
    dict_token["dict"] = dict;
    dict_token["access_token"] = token;
    dict_token["subdomain"] = subdomain;
    var settings = {
        url: server_url + '/zendesk/ticket_fields',
        type: 'POST',
        contentType: 'application/json',
        data: JSON.stringify(dict_token)
    };
    client.request(settings).then(
        function (data) {
            client.invoke('notify', 'Custom Ticket Field successfully added');
        showfinalWindow(client, token, server_url);
        },
        function (response) {
            var msg = 'Error ' + response.status;
            client.invoke('notify', msg, 'error');
        }
    );
    });
  }
  
  
  function showfinalWindow(client, token, server_url) {
    client.context().then(function(context) {
      subdomain = context.account.subdomain;
      dict_token = {};
      dict_token["access_token"] = token;
      dict_token["subdomain"] = subdomain;
      var settings = {
        url: server_url + '/zendesk/check_isactive',
        type: 'POST',
        contentType: 'application/json',
        data: JSON.stringify(dict_token)
      };
      client.request(settings).then(
        function (data) {
          showfinalWindow1(client, token, data, server_url)
        },
        function (response) {
          var msg = 'Error ' + response.status;
          client.invoke('notify', msg, 'error');
        }
      );
      });
  }
  
  
  function showfinalWindow1(client, token, flag, server_url) {
    var source = $("#final_window-hdbs").html();
    var template = Handlebars.compile(source);
    var html = template();
    $("#content").html(html);
  
    if (flag == 1){
      //document.getElementById("ticketStatus").innerHTML = '<strong>Ticket Tagging Active:</strong><br><span>All new tickets will be processed.</span>';
      document.getElementById("ticketStatus").innerHTML = '<strong>Ticket Tagging is Active:</strong><br>';    
      document.getElementById("onOff").checked = true;
    }
    else{
      //document.getElementById("ticketStatus").innerHTML = '<strong>Ticket Tagging InActive:</strong><br><span>All new tickets will not be processed.</span>';
      document.getElementById("ticketStatus").innerHTML = '<strong>Ticket Tagging is InActive:</strong><br>';
      document.getElementById("onOff").checked = false;
    }
  
    $("#previous_page").click(function (event) {
      event.preventDefault();
      showFieldModel(client, token, server_url);
    })
  
  
    $('.switch').click(function (event) {
      event.preventDefault();
      var log;
      const cb = document.getElementById('onOff');
      if (cb.checked == false){
        document.getElementById("onOff").checked = true;
        log = 1;
        document.getElementById("ticketStatus").innerHTML = '<strong>Ticket Tagging is Active:</strong><br>';
      }
      else{
        document.getElementById("onOff").checked = false;
        log = 0;
        document.getElementById("ticketStatus").innerHTML = '<strong>Ticket Tagging is InActive:</strong><br>';
      }
  
      client.context().then(function(context) {
        subdomain = context.account.subdomain;
    
        var settings = {
          url: server_url + '/zendesk/isactive',
          type: 'POST',
          contentType: 'application/json',
          data: JSON.stringify({"access_token":token, "subdomain":subdomain, "log":log})
        };
        client.request(settings).then(
          function (data) {          
          },
          function (response) {
            var msg = 'Error ' + response.status;
            client.invoke('notify', msg, 'error');
          }
          );
        });    
  
    });
  }
  
  
  function htmlToElement(html) {
    var template = document.createElement('template');
    html = html.trim(); // Never return a text node of whitespace as the result
    template.innerHTML = html;
    return template.content.firstChild;
  }
  