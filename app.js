
var Lang = require('./includes/lang.js');
var LogicBuilder = require('./includes/logic.js')(Lang);
var Scripter = require('./includes/psscript.js')(Lang, LogicBuilder);

// Predicates
var HasDeskPhone = new Lang.TransportRuleCondition('phonenumber', '-SenderADAttributeMatchesPatterns "phonenumber:\\w"');
var HasCellPhone = new Lang.TransportRuleCondition('mobilenumber', '-SenderADAttributeMatchesPatterns "mobilenumber:\\w"');
var HasVCard = new Lang.TransportRuleCondition('hasvcard', '-FromMemberOf "hasvcard"');
var HasInitials = new Lang.TransportRuleCondition('initials', '-SenderADAttributeMatchesPatterns "initials:\\w"');

function Not(predicate) {
    return new Lang.Negation(predicate);
}

// Common Card templates
var commonCards = {
    'A-1': [     HasDeskPhone ,     HasCellPhone ,     HasVCard , Not(HasInitials) ],
    'A-2': [     HasDeskPhone ,     HasCellPhone ,     HasVCard ,     HasInitials  ],
    'A-3': [     HasDeskPhone ,     HasCellPhone , Not(HasVCard)                   ],
    'B-1': [     HasDeskPhone , Not(HasCellPhone),     HasVCard , Not(HasInitials) ],
    'B-2': [     HasDeskPhone , Not(HasCellPhone),     HasVCard ,     HasInitials  ],
    'B-3': [     HasDeskPhone , Not(HasCellPhone), Not(HasVCard)                   ],
    'C-1': [ Not(HasDeskPhone),     HasCellPhone ,     HasVCard , Not(HasInitials) ],
    'C-2': [ Not(HasDeskPhone),     HasCellPhone ,     HasVCard ,     HasInitials  ],
    'C-3': [ Not(HasDeskPhone),     HasCellPhone , Not(HasVCard)                   ],
    'D-1': [ Not(HasDeskPhone), Not(HasCellPhone),     HasVCard , Not(HasInitials) ],
    'D-2': [ Not(HasDeskPhone), Not(HasCellPhone),     HasVCard ,      HasInitials ],
    'D-3': [ Not(HasDeskPhone), Not(HasCellPhone), Not(HasVCard)                   ],
};

// Generate the PowerShell Script
var structuredLogic = LogicBuilder.build(Lang.getAllPredicates(), commonCards);
process.stdout.write(Scripter.generate(Object.keys(commonCards), structuredLogic));
