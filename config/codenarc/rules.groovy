ruleset {
	ruleset('rulesets/basic.xml')
	ruleset('rulesets/braces.xml')
	ruleset('rulesets/concurrency.xml')
	ruleset('rulesets/convention.xml')
	ruleset('rulesets/design.xml')
	ruleset('rulesets/dry.xml') {
		exclude 'DuplicateNumberLiteral'
		exclude 'DuplicateMapLiteral'
		exclude 'DuplicateStringLiteral'
	}
	ruleset('rulesets/exceptions.xml') {
		exclude 'CatchException'
	}
	ruleset('rulesets/formatting.xml')
	ruleset('rulesets/generic.xml')
	ruleset('rulesets/groovyism.xml')
	ruleset('rulesets/imports.xml')
	ruleset('rulesets/jdbc.xml')
	ruleset('rulesets/junit.xml') {
		exclude 'JUnitSetUpCallsSuper'
	}
	ruleset('rulesets/logging.xml') {
		exclude 'Println'
	}
	ruleset('rulesets/naming.xml') {
		MethodName { regex = /[a-z][\w\s'\(\)#:]*/ }
		exclude 'FieldName'
		exclude 'PropertyName'
		FactoryMethodName { regex = /(build.*|make.*)/ }
	}	
	ruleset('rulesets/security.xml') {
		exclude 'JavaIoPackageAccess'
	}
	ruleset('rulesets/serialization.xml')
	ruleset('rulesets/size.xml') {
		CyclomaticComplexity { maxMethodComplexity = 30 }
	}
	ruleset('rulesets/unnecessary.xml') {
		exclude 'UnnecessaryCollectCall'
		exclude 'UnnecessarySubstring'
	}
	ruleset('rulesets/unused.xml')
}