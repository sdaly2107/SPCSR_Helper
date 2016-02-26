# SPCSR_Helper


Dealing with SharePoint Client Side rendering can be painful at times and time consuming having to create lots of boiler plate to perform simple tasks.  This library was created to lift some of that pain.  Mainly it provides the following - 

-CSR Template hooking instead of overriding - if you create an override you don't have to work out the correct default template to return, your hooks are passed the default template.  It's up to you if you want to modify the template or just do something else in the hook, for example register a validator.

-You no longer have to create an override just to get ctx.CurrentItem or ctx.ControlMode - just call SPCSR.CurrentItem or SPCSR.ControlMode to determine which form you are on.

-Field schema IDs are captured so elements on the UI can be found in a more robust manner using SPCSR.Utils.FindField()



TEMPLATES
**********

//PASS THROUGH FILTER ON ALL TEXT FIELDS, DEFAULT TEMPLATE WILL BE RETURNED
SPCSR.Utils.HookFieldTemplates({

        'Text': function(ctx, template) {

            console.log('Text field hook for ' + ctx.CurrentFieldSchema.Name);
        }

    }
);


//TEMPLATE OVERRIDE ON SPECIFIC FIELD
SPCSR.Utils.HookFieldTemplates({

        'MyField': function(ctx, template) {

			template.html = '<b>My Template</b>'
        }

    }
);


//TEMPLATE APPEND
SPCSR.Utils.HookFieldTemplates({

        'MyField': function(ctx, template) {

			template.html = template.html + ' <br /> Wow';
        }

    }
);


//ONLY ENABLE EDITING IN NEW FORM
SPCSR.Utils.HookFieldTemplates({

        'MyField': function(ctx, template) {

			template.enabled = ctx.ControlMode === SPClientTemplates.ClientControlMode.NewForm;
        }

    }
);


//DISABLING TEMPLATE WITHIN ASYNC CALL
SPCSR.Utils.HookFieldTemplates({

        'MyField': function(ctx, template) {

		template.defercallbacks = true; 

		SPCSR.Permissions.IsUserInGroups(groups).then(function (inGroup) {
				
			template.toggleEnabled(inGroup);
			template.registerCallbacks();
		});
        }

    }
);


//REGISTERING CONDITIONAL REQUIRED VALIDATOR - ONLY VALIDATE DATE FIELD IF NAME STARTS WITH MY
SPCSR.Utils.HookFieldTemplates({

        'DateTime': function(ctx, template) {

			var schema = ctx.CurrentFieldSchema;
			var conditionalValidator = new SPCSR.Validators.ConditionalRequiredValidator(function validationCondition() {

                		return schema.Name.indexOf('my') > -1;
            		});
			
			SPCSR.Utils.RegisterSingleValidator(ctx, conditionalValidator);
        }

    }
);


//MULTI HOOKS
SPCSR.Utils.HookFieldTemplates({

        'Text': function(ctx, template) {

            //do something
        },
	'User': function(ctx, template) {

            //do something
        },
	'MyField': function(ctx, template) {

            //do something
        }

    }
);



//SAME HOOK FOR MULTIPLE FIELDS
SPCSR.Utils.HookFieldTemplates({

        'FIELDA|FIELDB|FIELDC': function(ctx, template) {

            //do something
        }

    }
);


//HOOK ALL FIELDS
SPCSR.Utils.HookFieldTemplates({

        '*': function(ctx, template) {

            //do something
        }
    }
);



//HOOK ALL FIELDS AND ADD MORE SPECIFIC FILTER - EXECUTED IN ORDER OF DEFINITION
SPCSR.Utils.HookFieldTemplates({

        '*': function(ctx, template) {

            template.html = 'All text';
        },
	'MyField': function(ctx, template) {

            template.html = 'More specific template';
        }
    }
);



VALUE CAPTURE
*************

//ACCESS MyField VALUE ON RECORD

var fieldvalue = SPCSR.CurrentItem.MyField;


//DETECT IF DISPLAY/EDIT/NEW FORM

var displaymode = SPCSR.ControlMode;



FINDING FIELDS
**************

//FINDING FIELD FROM UI USING INTERNAL NAME AND SCHEMA ID FOR EXACT MATCH

var latestValue = SPCSR.Utils.FindField('MyField').val();



XML SPECIAL CHAR ENCODING
*************************


var internalName = SPCSR.Encoding.EncodeXMLSpecialChars('My Field?');
console.log(internalName); //gives 'My_x0020_Field_x003F_'



var decoded = SPCSR.Encoding.EncodeXMLSpecialChars(internalName);
console.log(decoded); //gives 'My Field?'



OTHERS
******

There's a few other little helpers in there too, for example email validator, functions to check if user is in groups etc. The main reason for this lib was for the CSR hook helper stuff though. 
