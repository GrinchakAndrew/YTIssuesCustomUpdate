var config = {
    transport: new(function() {
        var host = 'https://youtrack.oraclecorp.com/rest';
        this.getIssue = function(issue_id) {
            return $.ajax({
                url: host + '/issue/' + issue_id,
                dataType: "json",
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                }
            });
        };
        this.createIssue = function(data) {
            return $.ajax({
                url: host + '/issue',
                data: data,
                type: 'post',
                dataType: "json",
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                }
            });
        };
        this.updateIssue = function(issueId, data) {
            return $.ajax({
                url: host + '/issue/' + issueId,
                data: data,
                type: 'post',
                dataType: "json",
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                }
            });
        };
        this.execCommand = function(issueId, command) {
            var rawCommand = [];
            for (var key in command) {
                if (command[key].map && command[key].length > 1) {
                    rawCommand.push('add ' + key + ' ' + command[key].join(' '));
                } else if (key == 'Assignee') {
                    rawCommand.push(key + ' ' + command[key].join(key + ' '));
                } else {
                    rawCommand.push(key + ' ' + command[key]);
                }
            }
            return $.ajax({
                url: host + '/issue/' + issueId + '/execute',
                data: {
                    command: rawCommand.join(' ')
                },
                type: 'post',
                dataType: "json",
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                }
            });
        };
    })(),

    eventManager: new(function() {
        var pool = {};
        this.on = function(event, handler) {
            if (!pool[event]) {
                pool[event] = handler;
            }
        };
        this.off = function(event) {
            if (pool[event]) {
                delete pool[event];
            }

        };
        this.trigger = function(event, args) {
            if (pool[event] && typeof pool[event] === 'function') {
                return pool[event].apply(this, args);
            }
        };
    })(),
    sheetNames: [],
    wb: '',
    fnArr: [function(el) {
        $(el).css('background-color') !== "rgba(0, 0, 0, 0)" ?
            $(el).css('background-color', '') :
            $(el).css('background-color', '#CCEEFF');
    }],
    defPreventer: function(e) {
        e.originalEvent.stopPropagation();
        e.originalEvent.preventDefault();
        config.fnArr.forEach(function(i, j) {
            if (typeof i == 'function') {
                i(e.target);
            }
        });
        config.fnArr = [];
    },
    init: function(what) {
        what.forEach(function(el) {
            if ($(el).length) {
                $(el).on('dragover', config.defPreventer);
                $(el).on('dragenter', config.defPreventer);
            }
        });
    },
	
	rangeSeeker : function(workSheet /*Final List*/ , columnName /*Oracle Project Name*/ ) {
                    var workbook = config.wb['Workbook']['Sheets'];
                    var range;
                    var letterRanges = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
                    var ref;
                    var splitRefArrOf2;
                    var upperBoundNum;
                    var higherBoundNum;
                    var upperBoundLetter;
                    var lowerBoundLetter;
                    var columnNameLetter;
                    workbook.forEach(function(sheet) {
                        if (sheet['name'] == workSheet) {
                            ref = config.wb.Sheets[sheet['name']]['!ref'];
                            splitRefArrOf2 = ref.split(':');
                            upperBoundNum = parseInt(config.wb.Sheets[sheet['name']]['!ref'].split(':')[0].match(/\d+/));
                            lowerBoundNum = parseInt(config.wb.Sheets[sheet['name']]['!ref'].split(':')[1].match(/\d+/));
                            upperBoundLetter = ref.split(':')[0].match(/\D/)[0];
                            lowerBoundLetter = ref.split(':')[1].match(/\D/)[0];
                            for (var i = letterRanges.length; i--;) {
                                if (config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum] &&
                                    config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v']) {
                                    if (config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v'] == columnName ||
                                        config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v'].includes(columnName)) {
                                        range = letterRanges[i] + (upperBoundNum + 1) + ":" + letterRanges[i] + (upperBoundNum + 1);
                                    }
                                }
                            }
                        }
                    });
                    return range;
        },
	
    getItemNamesByColumn: function(workSheet, columnName) {
        var workbook = config.wb.Workbook.Sheets;
        if (config.wb.Sheets[workSheet]) {
            var keys = Object.keys(config.wb.Sheets[workSheet]); // issues
            var upperBound = parseInt(config.wb.Sheets[workSheet]['!ref'].split(':')[1].match(/\d+/));
            var returnable = [];
            var theKey = '';
			for (var i = 0; i < keys.length; i++) {
                        if (keys[i].match(/^[A-Z]+(\d+)/) && keys[i].match(/^[A-Z]+(\d+)/)[1] === '1') {
                            var _columnName =
                                config.wb.Sheets[workSheet][keys[i]] ?
                                config.wb.Sheets[workSheet][keys[i]]['v'] : '';
                            if (_columnName == columnName) {
                                theKey = keys[i];
                                break;
                            }
                        }
                    }
                    if (theKey) {
                        theKey = theKey.replace(/[0-9]+/, '');
                        while (upperBound > 1) {
                            config.wb.Sheets[workSheet][theKey + upperBound] &&
                                config.wb.Sheets[workSheet][theKey + upperBound]['v'] ?
                                returnable.push(config.wb.Sheets[workSheet][theKey + upperBound]['v']) : returnable;
                            upperBound--;
                        }
                   }
            return returnable.length ? returnable.reverse() : null;
        }
    },
    readFile: function(e) {
        if (e.originalEvent.dataTransfer) {
            if (e.originalEvent.dataTransfer.files.length) {
                var files = e.originalEvent.dataTransfer.files;
                config.f = files[0];
                var reader = new FileReader(),
                    name = config.f.name;
                reader.onload = function(e) {
                    var data = e.target.result;
                    config.wb = XLSX.read(data, {
                        type: 'binary'
                    });
                    if (!config.wb.SheetNames.some(function(sheet) {
                            if (~config.sheetNames.indexOf(sheet)) {
                                return true;
                            }
                        })) {
                        config.sheetNames = config.sheetNames.concat(config.wb.SheetNames);
                    }

                    if (!config.sheetNames.length) {
                        function UserException(message) {
                            this.message = message;
                            this.name = "UserException";
                        }
                        throw new UserException("The Excel File Seems To Have No Sheets!");
                        $('#drag-and-drop').addClass('failure');
                    }
                    config.eventManager.trigger('onFileRead');
                };
                reader.readAsBinaryString(config.f);
                config.fnArr.push(function(el) {
                    $(el).css('background-color') !== "rgba(0, 0, 0, 0)" ?
                        $(el).css('background-color', '') :
                        $(el).css('background-color', '#CCEEFF');
                });
                config.fnArr.forEach(function(i, j) {
                    if (typeof i == 'function') {
                        i(e.target);
                    }
                });
            }
        }
    }
};

$(document).ready(function() {
    config.init(['#draganddropitemsid']);
    $('#draganddropitemsid').on('drop',
        function(e) {
            config.defPreventer(e);
            config.readFile(e);
            config.eventManager.on('Issue Id', config.getItemNamesByColumn);
            config.eventManager.on('OA Name', config.getItemNamesByColumn);
            config.eventManager.on('Client Account', config.getItemNamesByColumn);
            config.eventManager.on('Campaign ID', config.getItemNamesByColumn);

            config.eventManager.on('onFileRead', function() {
                if (config.sheetNames.length) {
                    config.sheetNames.forEach(function(sheet) {
                        // the file with the unique field Issue Id , that is missing in the one generated by CSV file from YT, should be the first to grad-and-drop
						config[sheet] = config[sheet] ? config[sheet] : {};
						if(!config[sheet]['Issue Id'] && config.rangeSeeker(sheet, 'Issue Id')){
							config[sheet]['Issue Id'] = config.eventManager.trigger('Issue Id', [sheet, 'Issue Id']);
						}
						if(!config[sheet]['OA Name'] && config.rangeSeeker(sheet, 'OA Name')){
							config[sheet]['OA Name'] = config.eventManager.trigger('OA Name', [sheet, 'OA Name']);
						}
						if(!config[sheet]['Client Account'] && config.rangeSeeker(sheet, 'Client Account')){
							config[sheet]['Client Account'] = config.eventManager.trigger('Client Account', [sheet, 'Client Account']);
						}
						if(!config[sheet]['Campaign ID'] && config.rangeSeeker(sheet, 'Campaign ID')){
							config[sheet]['Campaign ID'] = config.eventManager.trigger('Campaign ID', [sheet, 'Campaign ID']);
						}
                    });
					config.eventManager.trigger('getItemNamesByColumn Done', []);
                } else {
                    throw new UserException("The Excel File Seems To Have No Sheets!");
                }
            });
			config.eventManager.on('getItemNamesByColumn Done', function() {
				if(config.sheetNames.every(function(sheet) {
					return config[sheet]['OA Name'] && config[sheet]['OA Name'].length && 
							config[sheet]['Client Account'] && config[sheet]['Client Account'].length
				}) && config.sheetNames.some(function(sheet){
					return config[sheet]['Campaign ID'] && config[sheet]['Campaign ID'].length
				}) && 
					config.sheetNames.some(function(sheet){
						return config[sheet]['Issue Id'] && config[sheet]['Issue Id'].length
				})) {
					config.eventManager.trigger('Reading All Complete', []); 	
				}
			});
        });
    config.eventManager.on('Reading All Complete', function() {
        console.log('Reading All Complete');
        $('#draganddropitemsid').addClass('success');
		var Ids1, ClientAcc1, OANames1, CampaignIDs, ClientAcc2, OANames2, Ids2Process = [];
		for(var i = 0; i < config.sheetNames.length; i++) {
			if(config[config.sheetNames[i]]) {
				if(~Object.keys(config[config.sheetNames[i]]).indexOf('Issue Id')){
					Ids1 = config[config.sheetNames[0]]['Issue Id'];
						if(~Object.keys(config[config.sheetNames[i]]).indexOf('Client Account')) {
							ClientAcc1 = config[config.sheetNames[i]]['Client Account'];
						}
						if(~Object.keys(config[config.sheetNames[i]]).indexOf('OA Name')) {
							OANames1 = config[config.sheetNames[i]]['OA Name'];
						}
				} else {
					if(~Object.keys(config[config.sheetNames[i]]).indexOf('Client Account')) {
							ClientAcc2 = config[config.sheetNames[i]]['Client Account'];
						}
					if(~Object.keys(config[config.sheetNames[i]]).indexOf('OA Name')) {
							OANames2 = config[config.sheetNames[i]]['OA Name'];
					}
					CampaignIDs = config[config.sheetNames[i]]['Campaign ID'];
				}
			}
		}
		if(Ids1 && Ids1.length) {
			for(var i = 0; i < Ids1.length; i++) {
				if(~ClientAcc2.indexOf(ClientAcc1[i]) && ~OANames2.indexOf(OANames1[i])){
					var pair = {};
					pair['id'] = Ids1[i];
					pair['writable'] = CampaignIDs[OANames2.indexOf(OANames1[i])];
					Ids2Process.push(pair);
				}
			}
		}
		if(Ids2Process && Ids2Process.length){
			config['Ids2Process'] = Ids2Process;
		}
		
		config['Ids2Process'].forEach(function(pair) {
			var id = pair['id'];
			var writable = pair['writable'];
			config.eventManager.on(id, [id, writable]);
			config.eventManager.trigger('transport do', [id, writable]);
		});
		
    }); // the end of 'Reading All Complete' line
	config.eventManager.on('transport do', function(id, writable) {
			if(id && writable) {
				config.transport.getIssue(id).done(function(data){
					console.log(data)
				}).done(function() {
					config.transport.execCommand(id, {'Campaign ID' : writable});
				});
			}
		});
		
});