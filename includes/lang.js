module.exports = (function () {
    var allPredicateList = [];
    function Predicate() {
        allPredicateList.push(this);
    }
    Predicate.prototype.getId = function () {
        return allPredicateList.indexOf(this);
    };
    Predicate.prototype.getPredicate = function () {
        return this;
    };

    function TransportRuleCondition(name, flags) {
        this.name = name;
        this.flags = flags;
        allPredicateList.push(this);
    }
    TransportRuleCondition.prototype = new Predicate;
    TransportRuleCondition.prototype.constructor = TransportRuleCondition;
    TransportRuleCondition.constructor = Predicate;

    function Negation(condition) {
        if (!(condition instanceof TransportRuleCondition)) {
            throw 'Expected TransportRuleCondition.';
        }
        this.condition = condition;
    }
    Negation.prototype = new Predicate;
    Negation.prototype.constructor = Negation;
    Negation.constructor = Predicate;
    Negation.prototype.getPredicate = function () {
        return this.condition;
    };

    return {
        getAllPredicates: function () {
            return allPredicateList.slice(0);
        },
        TransportRuleCondition: TransportRuleCondition,
        Negation: Negation,
    };
}());
