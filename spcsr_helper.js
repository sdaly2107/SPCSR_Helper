'use strict'

window.console && console.log('SPCSR loaded...');

var SPCSR = SPCSR || {};
SPCSR.Validators = SPCSR.Validators || {};

SPCSR.CurrentItem = null;
SPCSR.ControlMode = null;

SPCSR.Validators.ConditionalRequiredValidator = function(fnIsRequired) {

    this._fnIsRequired = fnIsRequired;

    SPCSR.Validators.ConditionalRequiredValidator.prototype.ValidateField = function(value) {

        if (this._fnIsRequired()) {

            var requiredValidator = new SPClientForms.ClientValidation.RequiredValidator();
            return requiredValidator.Validate(value);
        }

        return new SPClientForms.ClientValidation.ValidationResult(false, null); //valid result
    };

};

SPCSR.Validators.EmailValidator = function(onlyValidateIfFilled) {

    this._onlyValidateIfFilled = onlyValidateIfFilled;

    function isValidEmail(emailAddress) {

        var emailPattern = new RegExp(/^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?$/i);

        return emailPattern.test(emailAddress);
    };

    SPCSR.Validators.EmailValidator.prototype.ValidateField = function(value) {

        var isSet = typeof value === 'string' && value.length > 0;

        var isValid = isValidEmail(value);

        if (this._onlyValidateIfFilled && !isValid && !isSet) {

            isValid = true;
        }

        return new SPClientForms.ClientValidation.ValidationResult(!isValid, 'Enter a valid email address'); //valid result
    };

};

SPCSR.Encoding = (function() {

    //TODO: Complete with all from https://www.w3.org/Style/XSL/TestSuite/results/4/XEP/symbol.pdf
    var mappings = {
        'x005F': '_',
		'x0020': ' ',
        'x0021': '!',
        'x0023': '#',
        'x0025': '%',
        'x0026': '&',
        'x0028': '(',
        'x0029': ')',
        'x002B': '+',
        'x002C': ',',
        'x002E': '.',
        'x002F': '/',
        'x003F': '?',
        'x003A': ':',
        'x003B': ';',
        'x003C': '<',
        'x003D': '=',
        'x003E': '>',
        'x005B': '[',
        'x005D': ']',        
        'x007B': '{',
        'x007C': '|',
        'x007D': '}'
    };

    var pattern = new RegExp('_x[a-zA-Z0-9]*_');

    return {
        DecodeXMLSpecialChars: decodeXMLSpecialChars,
        EncodeXMLSpecialChars: encodeXMLSpecialChars
    };

    function decodeXMLSpecialChars(input) {

        if (pattern.test(input)) {

            for (var kMapping in mappings) {

                input = input.replace(new RegExp('_' + kMapping + '_', 'g'), mappings[kMapping]);
            }
        }

        return input;
    }

    function encodeXMLSpecialChars(input) {

        for (var kMapping in mappings) {

            var replacement = escapeRegExp(mappings[kMapping]);

            input = input.replace(new RegExp(replacement, 'g'), '_' + kMapping + '_');
        }

        return input;
    }

    function escapeRegExp(input) {

        return input.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
    }


})();


SPCSR.Permissions = (function() {
	
	var _groupsCache = null;
	var _ctx = null;
		
    return {

        GetCurrentUser: getCurrentUser,
        IsUserInGroups: isUserInGroups
    };

    function getContext() {

        var deferred = jQuery.Deferred();

		if(null !== _ctx){
			
			deferred.resolve(_ctx);
		}
		
        ExecuteOrDelayUntilScriptLoaded(function() {

            _ctx = SP.ClientContext.get_current();
            deferred.resolve(_ctx);

        }.bind(this), 'sp.js');

        return deferred.promise();
    }

    function getCurrentUser() {

        var deferred = jQuery.Deferred();

        jQuery.when(getContext()).then(function(context) {


            var currentUser = context.get_web()
                .get_currentUser();

            context.load(currentUser);

            context.executeQueryAsync(
                function onUserLoaded() {

                    deferred.resolve(currentUser);
                },
                function onUserLoadFail(sender, args) {

                    deferred.reject(sender, args);
                }
            );

        });

        return deferred.promise();
    }


    function isUserInGroups(names) {
		
        if (typeof names === 'string') {

            names = [names];
        }
		
		var _names = names;
		

        var deferred = jQuery.Deferred();

        if (null === _groupsCache) {

             jQuery.when(getContext()).then(function(context){
				
                var currentUser = context.get_web().get_currentUser();
                var groups = currentUser.get_groups();
                context.load(currentUser);
                context.load(groups);
                context.executeQueryAsync(
                    function onGroupsLoaded(sender, args) {
					
						if(null !== _groupsCache){
							
							 deferred.resolve(checkIfUserIsInGroup());
						}
						
						var groupNames = [];
                        var e = groups.getEnumerator();
                        while (e.moveNext()) {

                            groupNames.push(e.get_current().get_title());
                        }
						
						_groupsCache = groupNames;

                        deferred.resolve(checkIfUserIsInGroup());

                    }.bind(this),
                    function OnGroupsLoadFailure(sender, args) {

                        deferred.fail(sender, args);
                    }
                );

            }.bind(this));

        } else {
             deferred.resolve(checkIfUserIsInGroup()); //groups cached so check now
        }

        function checkIfUserIsInGroup() {

            for (var groupIndex = 0; groupIndex < _groupsCache.length; ++groupIndex) {

                var groupTitle = _groupsCache[groupIndex];
                if (_names.indexOf(groupTitle) !== -1) {

                    return true;
                }
            }

           return false;
        }

        return deferred.promise();
    }

})();


SPCSR.Utils = (function() {

    var _hooks;
    var _fieldIDs = {};
    var displayNames = ['Invalid', 'DisplayForm', 'EditForm', 'NewForm', 'View'];

    var defaultTemplates = {
        Templates: {
            Fields: SPClientTemplates._defaultTemplates.Fields.default.all.all
        }
    }

    return {
        HookFieldTemplates: hookFieldTemplates,
        RegisterSingleValidator: registerSingleValidator,
        FindField: findField
    };


    function registerSingleValidator(ctx, validator) {

        var validators = new SPClientForms.ClientValidation.ValidatorSet();
        validators.RegisterValidator(validator);

        ctx.FormContext.registerClientValidator(ctx.CurrentFieldSchema.Name, validator);
    }

    function findField(internalName, throwIfNotFound) {

        if (typeof throwIfNotFound === 'undefined') {

            throwIfNotFound = true;
        }

        if (!_fieldIDs.hasOwnProperty(internalName)) {

            throw 'No controlID captured for ' + internalName;
        }

        var partialControlID = _fieldIDs[internalName];
        partialControlID = escapeSelector(partialControlID);

        var field = jQuery("[id^='" + partialControlID + "']");
        if (throwIfNotFound && 0 === field.length) {

            throw 'Field for ' + internalName + ' was not found';
        }

        return field;
    }

    function escapeSelector(selector) {

        if (selector) {

            return selector.replace(/([ #;?%&,.+*~\':"!^$[\]()=>|\/@])/g, '\\$1');
        }

        return selector;
    }

    function deepCopy(object) {

        var copy = {};
        jQuery.extend(true, copy, object);

        return copy;
    }

    function getExistingTemplateOverrides() {

        var existingOverrides;
        if (SPClientTemplates.TemplateManager._TemplateOverrides.Fields.default) {

            existingOverrides = deepCopy(SPClientTemplates.TemplateManager._TemplateOverrides.Fields.default.all.all);
        }

        return existingOverrides;
    }

    function createHook(originalCall) {

        var hookedCall = function SPCSRHookTemplate(ctx) {

            var html = templateFixes(ctx, originalCall(ctx)); //originalCall may be default or custom template 

            return hook(ctx, html);
        };

        return hookedCall;
    }

    function hookFieldTemplates(hooks) {

        _hooks = hooks;

        var existingOverrides = getExistingTemplateOverrides();
        var newOverrides = {};

        //merge existing templates with new default templates
        jQuery.extend(true, newOverrides, defaultTemplates.Templates.Fields, existingOverrides);

        //append hook function to each default or custom template
        for (var kTemplate in newOverrides) {

            for (var kDisplayTemplate in newOverrides[kTemplate]) {

                var originalCall = newOverrides[kTemplate][kDisplayTemplate];
                var hookedCall = createHook(originalCall);

                newOverrides[kTemplate][kDisplayTemplate] = hookedCall;
            }
        }

        var hookedTemplates = {
            Templates: {
                Fields: newOverrides
            }
        };

        // window.console && console.log(hookedTemplates);

        SPClientTemplates.TemplateManager.RegisterTemplateOverrides(hookedTemplates);
    }

    function captureCurrentItem(ctx) {

        if (null === SPCSR.CurrentItem && ctx.CurrentItem) {

            SPCSR.CurrentItem = deepCopy(ctx.CurrentItem);
        }

        if (null === SPCSR.ControlMode) {

            SPCSR.ControlMode = ctx.ControlMode;
        }
    }

    function captureControlID(ctx) {

        var partialFieldID = ctx.CurrentFieldSchema.Name + '_' + ctx.CurrentFieldSchema.Id + '_';
        _fieldIDs[ctx.CurrentFieldSchema.Name] = partialFieldID;
    }

    function isInternalField(name) {

        return name === 'Author' ||
            name === 'Editor' ||
            name === '_UIVersionString' ||
            name === 'Created' ||
            name === 'Attachments' ||
            name === 'Modified';
    }

    function templateFixes(ctx, html) {

        if (ctx.CurrentFieldSchema.FieldType === 'Note') {

            //if disabling template then enabling, note field does not show with border, fix css
            html = html.replace('ms-rtestate-field ms-rtefield', 'ms-rtestate-field ms-rtefield ms-inputBox ms-inputBoxActive');
        }

        return html;
    }


    function hook(ctx, html) {

        var fieldInternalName = ctx.CurrentFieldSchema.Name;
        var fieldType = ctx.CurrentFieldSchema.FieldType;

        captureCurrentItem(ctx);
        captureControlID(ctx);

        var displayViewName = displayNames[ctx.ControlMode];

        //ensure SP fields are readonly in edit mode
        if (ctx.ControlMode === SPClientTemplates.ClientControlMode.EditForm && isInternalField(fieldInternalName)) {
            displayViewName = displayNames[SPClientTemplates.ClientControlMode.DisplayForm];
        }

        var ctxCopy = deepCopy(ctx); //important, getReadOnlyTemplate prepares and updates some values on ctx, so we must copy
        var defaultDisplayTemplate = getReadOnlyTemplate(ctxCopy);

        if (ctx.ControlMode === SPClientTemplates.ClientControlMode.DisplayForm) {

            html = wrapWithSpanIdentifier(html, ctx); //add span with ID of name_schemaID_ so elements can be found on display page
        }

        var fieldElement;
        var template = {
            html: html,
            enabled: true, //set to false to always use a display template - cannot be used in async callbacks since field will be rendered already
            defaultHtml: html, //hold onto original templates incase user toggles between enabled and disabled states
            defaultDisplayHtml: defaultDisplayTemplate,
            update: function(html) { //allows updating of field html after rendering

                findField(fieldInternalName).replaceWith(html);
            },
            disable: function() {

                this.toggleEnabled(false);
            },
            enable: function() {

                this.toggleEnabled(true);
            },
            toggleEnabled: function(enabled) { //can be used in callbacks - the field is found and updated

                if (this.enabled === enabled) {

                    return;
                }

                this.enabled = enabled;

                var field = findField(fieldInternalName, false);

                if (0 === field.length) { //field wasn't found, most likely not rendered yet so set enabled to false - display template should then be used to render

                    return;
                }

                if (enabled) {

                    field.replaceWith(this.defaultHtml);
                } else {

                    field.replaceWith(this.defaultDisplayHtml);
                }

            },
            defercallbacks: false, //once a getvalue callback has been set, it cannot be unset
            registerCallbacks: function() {

                if (!this.enabled) {

                    if (!isInternalField(fieldInternalName)) { //don't mess around with fields such as attachments, created etc


                        ctx.FormContext.registerGetValueCallback(fieldInternalName, function() { //SP gets values from DOM to save, since we are using a display template we have to provide the values

                            if (fieldType === 'User' || fieldType === 'UserMulti') { //SP people picker template gets value to save from hidden field, we need to build this json up ourself

                                var userString = buildUserValueString(SPCSR.CurrentItem[fieldInternalName]);

                                return userString;
                            }


                            return SPCSR.CurrentItem[fieldInternalName]; //chuck back current value
                        });

                        //since we use a display template, there's no error span to show any validation errors.  in theory there should be none, but if something goes wrong then log it
                        ctx.FormContext.registerValidationErrorCallback(fieldInternalName, function(errorResult) {

                            if (errorResult.hasOwnProperty('validationError') && errorResult.validationError) {

                                console.error('Validation error on ' + fieldInternalName + ' = ' + errorResult.errorMessage);
                            }

                        });

                        //original peoplepicker callback calls EnsurePeoplePickerScript(InitControl) which does some dom manipulation and fails because elements don't exist if using display template.  if the template is disabled after rendering, for example from an async call then this won't take effect, because the init callback will have already run.
                        if (fieldType === 'User' || fieldType === 'UserMulti') {

                            ctx.FormContext.registerInitCallback(fieldInternalName, function() {

                            });
                        }


                    }

                }
            }
        };

        var hooks = getHooks(ctx); //multi hooks can be specified and are executed in order of definition, example a field could match for '*', 'Text' and 'MyFieldName'
        var hooksLength = hooks.length;
        for (var hookIndex = 0; hookIndex < hooksLength; ++hookIndex) {

            var fnHook = hooks[hookIndex];
            if (fnHook) {

                fnHook(ctx, template);
            }
        }

        if (!template.enabled) {

            template.html = defaultDisplayTemplate; //use display template
        }

        if (!template.defercallbacks) {

            template.registerCallbacks();
        }

        logContextErrors(ctx);


        return template.html;
    }

    function buildUserValueString(userFieldValue) {

        var itemObjs = [];

        for (var kItem in userFieldValue) {

            var userObject = userFieldValue[kItem];

            if (typeof userObject === 'object') {

                itemObjs.push({
                    "Key": userObject.Key,
                    "Description": userObject.Description,
                    "DisplayText": userObject.DisplayText,
                    "EntityType": userObject.EntityData.PrincipalType,
                    "ProviderDisplayName": userObject.ProviderDisplayName,
                    "ProviderName": userObject.ProviderName,
                    "IsResolved": userObject.IsResolved,
                    "EntityData": userObject.EntityData,
                    "MultipleMatches": [],
                    "AutoFillKey": userObject.Key,
                    "AutoFillDisplayText": userObject.DisplayText,
                    "AutoFillSubDisplayText": "",
                    "DomainText": window.location.hostname,
                    "Resolved": userObject.IsResolved
                });
            }

        }

        var userString = JSON.stringify(itemObjs);

        return userString;
    }

    function logContextErrors(ctx) {

        if (ctx.hasOwnProperty('Errors') && ctx.Errors.length > 0) { //SP catches all errors during template execution, output these so we know about them

            window.console && console.error(ctx.Errors);
        }
    }

    function getHooks(ctx) {

        var hooks = [];
        for (var kHook in _hooks) {

            if (kHook === '*' || kHook === ctx.CurrentFieldSchema.Name || kHook === ctx.CurrentFieldSchema.FieldType) {

                hooks.push(_hooks[kHook]);

            } else if (kHook.indexOf('|') !== -1) { //support hooks for multi field names, example MyFieldA|MyFieldB

                var multiNames = kHook.split('|');

                for (var nameIndex = 0; nameIndex < multiNames.length; ++nameIndex) {

                    if (ctx.CurrentFieldSchema.Name === multiNames[nameIndex]) {

                        hooks.push(_hooks[kHook]);
                    }
                }
            }
        }

        return hooks;
    }

    function wrapWithSpanIdentifier(html, ctx) {

        var id = ctx.CurrentFieldSchema.Name + '_' + ctx.CurrentFieldSchema.Id + '_$customhandle';

        html = '<span id="' + id + '">' + html + '</span>';

        return html;
    }


    function prepareField(ctx) {

        var fieldtype = ctx.CurrentFieldSchema.FieldType;

        if (fieldtype === 'User' || fieldtype === 'UserMulti') {

            prepareUserFieldValue(ctx);

        } else if (fieldtype === 'MultiChoice') {

            prepareMultiChoiceFieldValue(ctx);

        } else if (fieldtype === 'Note') {

            prepareNoteFieldValue(ctx);
        }
    }

    function getReadOnlyTemplate(ctx) {

        var displayMode = SPClientTemplates.ClientControlMode.DisplayForm;

        prepareField(ctx);

        var fnDefaultDisplayTemplate = defaultTemplates.Templates.Fields[ctx.CurrentFieldSchema.FieldType][displayNames[displayMode]];

        var html = fnDefaultDisplayTemplate(ctx);

        //in non display mode, fields have an ID made up of name_schemID_type, in display mode there is no identifier so we add one here
        html = wrapWithSpanIdentifier(html, ctx);

        return html;
    }

    function prepareUserFieldValue(ctx) {

        var userField = ctx.CurrentItem[ctx.CurrentFieldSchema.Name];

        if (typeof userField === 'string') {

            return; //value is already in correct format for display template
        }

        var fieldValue = '';

        for (var i = 0; i < userField.length; i++) {
            fieldValue += userField[i].EntityData.SPUserID + SPClientTemplates.Utility.UserLookupDelimitString + userField[i].DisplayText;

            if ((i + 1) != userField.length) {
                fieldValue += SPClientTemplates.Utility.UserLookupDelimitString
            }
        }

        ctx['CurrentFieldValue'] = fieldValue;
    }

    function prepareMultiChoiceFieldValue(ctx) {

        if (ctx['CurrentFieldValue']) {

            var fieldValue = ctx['CurrentFieldValue'];

            var find = ';#';
            var regExpObj = new RegExp(find, 'g');

            fieldValue = fieldValue.replace(regExpObj, '; ');
            fieldValue = fieldValue.replace(/^; /g, '');
            fieldValue = fieldValue.replace(/; $/g, '');

            ctx['CurrentFieldValue'] = fieldValue;
        }
    }

    function prepareNoteFieldValue(ctx) {

        if (ctx['CurrentFieldValue']) {

            var fieldValue = ctx['CurrentFieldValue'];
            fieldValue = "<div>" + fieldValue.replace(/\n/g, '<br />'); + "</div>";

            ctx['CurrentFieldValue'] = fieldValue;
        }
    }


})();
