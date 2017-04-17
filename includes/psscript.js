module.exports = function (Lang, LogicBuilder) {
    function Rule(level, desc, args) {
        this.name = (new Array(level + 2)).join('>') + ' ' + desc;
        this.args = args;
    }

    function getScopeArgs(level) {
        return (new Array(level)).fill(null).map(function (_, i) {
            return 'CC-Scope-' + (i + 1);
        }).join(' ');
    }

    function getDisclaimerFlags(key) {
        return '-ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["' + key + '"] -ApplyHtmlDisclaimerFallbackAction Wrap';
    }

    function translateToRules(level, structuredLogic) {
        var psscript = [];

        var scopeFlags = '-HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "' + (level > 0 ? 'CC-Scope-' + level : 'Ready') + '"';
        var elseFlags = '-ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-' + (level + 1) + '"';
        var enterScopeAction = '-SetHeaderName "X-SignatureCards" -SetHeaderValue "' + getScopeArgs(level + 1) + '"';
        var finishedAction = '-SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"';

        for (var i = 0; i < structuredLogic.length; i++) {
            var block = structuredLogic[i];
            if (!(block instanceof LogicBuilder.IfElse)) {
                var disclaimerFlags = getDisclaimerFlags(block);
                psscript.push(new Rule(
                    level, block,
                    [scopeFlags, disclaimerFlags, finishedAction]
                ));
            } else {
                // Consequent
                var condition = block.condition.flags;
                if (block.consequent.length === 1 && !(block.consequent[0] instanceof LogicBuilder.IfElse)) {
                    var disclaimerFlags = getDisclaimerFlags(block.consequent[0]);
                    psscript.push(new Rule(
                        level, 'IF ' + block.condition.name + ' THEN ' + block.consequent[0],
                        [condition, scopeFlags, disclaimerFlags, finishedAction]
                    ));
                } else {
                    psscript.push(new Rule(
                        level, 'IF ' + block.condition.name,
                        [condition, scopeFlags, enterScopeAction]
                    ));
                    psscript = psscript.concat(translateToRules(level + 1, block.consequent));
                }

                // Alternative
                if (block.alternative.length === 1 && !(block.alternative[0] instanceof LogicBuilder.IfElse)) {
                    var disclaimerFlags = getDisclaimerFlags(block.alternative[0]);
                    psscript.push(new Rule(
                        level, 'ELSE ' + block.alternative[0],
                        [elseFlags, scopeFlags, disclaimerFlags, finishedAction]
                    ));
                } else {
                    psscript.push(new Rule(
                        level, 'ELSE',
                        [elseFlags, scopeFlags, enterScopeAction]
                    ));
                    psscript = psscript.concat(translateToRules(level + 1, block.alternative));
                }
            }
        }
        return psscript;
    }

    function retrieveCommonCardTemplates(commonCardKeys) {
        var pscmd_templates = [ '$CommonCardTemplates = @{' ].concat(commonCardKeys.map(function (key) {
            return '    "' + key + '" = Get-Content $PSScriptRoot"\\' + key + '.txt";';
        }));
        pscmd_templates.push('}');

        var pscmd_verify = [
            '$CommonCardTemplatesNK = ($CommonCardTemplates.Keys | Measure-Object).Count',
            '$CommonCardTemplatesNV = ($CommonCardTemplates.Values | Measure-Object).Count',
            'if (-not $CommonCardTemplatesNK -eq $CommonCardTemplatesNV ) {',
            '    Write-Error "At least one Common Card template file could not be found or accessed in this folder. Please make sure it exists and then try again. Aborting."',
            '    exit',
            '}',
        ].join('\r\n');

        var pscmd_stringify = [ '$CommonCardTemplates = @{' ].concat(commonCardKeys.map(function (key) {
            return '    "' + key + '" = $CommonCardTemplates["' + key + '"] | Out-String;';
        }));
        pscmd_stringify.push('}');

        return pscmd_templates.join('\r\n') + '\r\n\r\n' + pscmd_verify + '\r\n\r\n' + pscmd_stringify.join('\r\n');
    }

    function retrieveCustomCardTemplates() {
        var pscmd_templates = [ '$CustomCardTemplates = @{}' ];
        pscmd_templates.push('ForEach ($CustomCardFile in Get-ChildItem $PSScriptRoot -Filter "*@*.txt") {');
        pscmd_templates.push('    $CustomCardTemplates[$CustomCardFile.BaseName] = Get-Content $CustomCardFile.VersionInfo.FileName | Out-String');
        pscmd_templates.push('}');

        var pscmd_verify = [
            '$CustomCardTemplatesNK = ($CustomCardTemplates.Keys | Measure-Object).Count',
            '$CustomCardTemplatesNV = ($CustomCardTemplates.Values | Measure-Object).Count',
            'if (-not $CustomCardTemplatesNK -eq $CustomCardTemplatesNV ) {',
            '    Write-Error "At least one Custom Card template file could not be read in this folder. Please check permissions and then try again. Aborting."',
            '    exit',
            '}',
        ].join('\r\n');
        return pscmd_templates.join('\r\n') + '\r\n\r\n' + pscmd_verify;
    }

    function generateCommonCardRules(structuredLogic) {
        var rules = translateToRules(0, structuredLogic);
        var count = 0;
        return rules.map(function (rule) {
            return ['New-TransportRule -Name "Signature | CC-' + (++count) + ' | ' + rule.name + '"'].concat(rule.args).join(' `\r\n ');
        }).join('\r\n\r\n');
    }

    function generateCustomCardRules() {
        var pscmd = [ 'ForEach ($EmailAddress in $CustomCardTemplates.Keys) {' ];
        pscmd.push('    $Disclaimer = $CustomCardTemplates[$EmailAddress]');

        var psrule = [
            '    New-TransportRule -Name "Signature | Custom Card | $($EmailAddress)"',
            '    -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "Ready"',
            '    -From $EmailAddress',
            '    -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $Disclaimer -ApplyHtmlDisclaimerFallbackAction Wrap',
            '    -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"'
        ].join(' `\r\n ');
        pscmd.push(psrule);

        pscmd.push('}');
        return pscmd.join('\r\n');
    }

    function generate(commonCardKeys, structuredLogicForCommonCards) {
        var psscript = [];
        psscript.push([
            'if (!$PSScriptRoot) {',
            '    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent',
            '}',
        ].join('\r\n'));

        psscript.push('Write-Output "Retrieving signature card templates..."');
        psscript.push(retrieveCommonCardTemplates(commonCardKeys));
        psscript.push(retrieveCustomCardTemplates());

        psscript.push('Write-Output "Uninstalling old signature card templates..."');
        psscript.push('$OldTransportRules = Get-TransportRule "Signature | *"');
        psscript.push('$OldTransportRules | %{ Remove-TransportRule -Identity $_.Guid.Guid -Confirm:$false }');

        psscript.push('Write-Output "Installing new signature card templates..."');
        psscript.push('New-TransportRule -Name "Signature | Reset" -RemoveHeader "X-SignatureCards"');

        var pscmd = [ 'New-TransportRule -Name "Signature | Start" -Enabled $false' ];
        pscmd.push('-SentToScope "NotInOrganization" -ExceptIfHeaderMatchesMessageHeader "In-Reply-To" -ExceptIfHeaderMatchesPatterns "\\w"');
        pscmd.push('-SetHeaderName "X-SignatureCards" -SetHeaderValue "Ready"');
        psscript.push(pscmd.join(' `\r\n '));

        psscript.push(generateCustomCardRules());
        psscript.push(generateCommonCardRules(structuredLogicForCommonCards));
        psscript.push('New-TransportRule -Name "Signature | End" -RemoveHeader "X-SignatureCards"');

        psscript.push('Write-Output "Enabling new signature card templates..."');
        psscript.push('Enable-TransportRule -Identity "Signature | Start"');
        psscript.push('Write-Output "Done."');

        return psscript.join('\r\n\r\n') + '\r\n';
    }

    return {
        generate: generate,
    };
};
