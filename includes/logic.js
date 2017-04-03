module.exports = function (Lang) {
    function IfElse(condition, consequent, alternative) {
        this.condition = condition;
        this.consequent = consequent;
        this.alternative = alternative;
    }

    function build(predicateList, cards) {
        if (predicateList.length === 0) {
            // No more predicates to check; return the matching card keys.
            return Object.keys(cards);
        }

        var ifBranch = {};
        var elseBranch = {};
        var other = {};

        for (var key in cards) {
            // Find matching predicates in this card.
            var matches = cards[key].filter(function (item) {
                return item.getPredicate() === predicateList[0];
            });

            // Move card into the appropriate bucket.
            switch (matches.length) {
            case 1:
                if (matches[0] instanceof Lang.Negation) {
                    elseBranch[key] = cards[key];
                } else {
                    ifBranch[key] = cards[key];
                }
                break;
            case 0:
                other[key] = cards[key];
                break;
            default:
                throw 'Binary predicate is used twice in the same template (' + key + ').';
            }
        }

        var result = [];

        // Recurse
        var consequent = build(predicateList.slice(1), ifBranch);
        var alternative = build(predicateList.slice(1), elseBranch);
        if (consequent.length > 0 || alternative.length > 0) {
            result.push(new IfElse(predicateList[0], consequent, alternative));
        }

        // Continuation
        var otherResult = build(predicateList.slice(1), other);
        result = result.concat(otherResult);

        return result;
    }

    return {
        build: build,
        IfElse: IfElse,
    }
};
