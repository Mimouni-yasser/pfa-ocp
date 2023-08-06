const user_input = $(".user-input")
const IP_div = $('#replaceable-content')
const endpoint = '/'
const delay_by_in_ms = 200
let scheduled_function = false
let last_IP_search


function isEqual(obj1, obj2) {
    var props1 = Object.getOwnPropertyNames(obj1);
    var props2 = Object.getOwnPropertyNames(obj2);
    if (props1.length != props2.length) {
        return false; 
    }
    for (var i = 0; i < props1.length; i++) {
        let val1 = obj1[props1[i]];
        let val2 = obj2[props1[i]];
        let isObjects = isObject(val1) && isObject(val2);
        if (isObjects && !isEqual(val1, val2) || !isObjects && val1 !== val2) {
            return false;
        }
    }
    return true;
}
function isObject(object) {
    return object != null && typeof object === 'object';
}

var changed = function(instance, cell, x, y, value) {
    clearInterval(scheduled_function)
    IP = results_sheet.getCellFromCoords(0, y).innerHTML
    MAC = results_sheet.getCellFromCoords(1, y).innerHTML
    COMMENT = results_sheet.getCellFromCoords(2,y).innerHTML
    TYPE = results_sheet.getCellFromCoords(3,y).innerHTML
    
    $.post('update/', {ip: IP, mac:MAC, comment: COMMENT, type: TYPE})
        .done(function(kk){
            console.log(kk)
            scheduled_function = setInterval(() => {
                ajax_call(endpoint, last_IP_search)
                }, 2000);
        })
}

results_sheet = jspreadsheet(document.getElementById('spreadsheet'), {
    columns: [

        { type: 'text', title:'IP', width:200 },
        { type: 'text', title:'MAC', width:200 },
        { type: 'text', title:'COMMENTAIRE', width:300 },
        { type: 'text', title:'type', width:200 },
        { type: 'calendar', title:'date ajouter', width:200 },
    ],
    allowInsertColumn: false,
    allowDeleteRow: false,
    allowDeleteColumn: false,
    allowInsertRow: false,
    onchange: changed,
});

let ajax_call = function (endpoint, request_parameters) {
    $.getJSON(endpoint, request_parameters)
        .done(response => {
                sheet_data = results_sheet.getData()
                data = JSON.parse(response)
                data.forEach((e,i,arr)=>arr[i]=Object.values(e['fields']),data)
                if(isEqual(sheet_data, data) === false)
                    results_sheet.setData(data)

                })
}


user_input.on('keyup', function () {
    
    const request_parameters = {
        ip: $('#IP-search').val(), // value of user_input: the HTML element with ID user-input
        mac: $('#MAC-search').val(),
        comment: $('#comment-search').val()
    }
    
    last_IP_search = request_parameters

    ajax_call(endpoint, request_parameters)
})

$(document).ready(function()
{
    const request_parameters = {
        q: ''
    }
    ajax_call(endpoint, request_parameters)
})

scheduled_function = setInterval(() => {
ajax_call(endpoint, last_IP_search)
}, 2000);


