ruleset {
	ruleset('rulesets/basic.xml')
	ruleset('rulesets/braces.xml')
	ruleset('rulesets/concurrency.xml')
	ruleset('rulesets/convention.xml')
	ruleset('rulesets/design.xml')
	ruleset('rulesets/dry.xml') {
		exclude 'DuplicateNumberLiteral'
		exclude 'DuplicateListLiteral'
		exclude 'DuplicateMapLiteral'
		exclude 'DuplicateStringLiteral'
	}
	ruleset('rulesets/exceptions.xml')
	ruleset('rulesets/formatting.xml')
	ruleset('rulesets/generic.xml')
	ruleset('rulesets/grails.xml')
	ruleset('rulesets/groovyism.xml')
	ruleset('rulesets/imports.xml')
	ruleset('rulesets/jdbc.xml')
	ruleset('rulesets/junit.xml') {
		exclude 'JUnitStyleAssertions'
		exclude 'JUnitSetUpCallsSuper'
	}
	ruleset('rulesets/logging.xml')
	ruleset('rulesets/naming.xml') {
		MethodName {
			regex = /[a-z][\w\s'\(\)]*/
		}
	}	
	ruleset('rulesets/security.xml')
	ruleset('rulesets/serialization.xml')
	ruleset('rulesets/size.xml')
	ruleset('rulesets/unnecessary.xml')
	ruleset('rulesets/unused.xml')
}