export interface Flags {
	Flag: string[];
}

export interface RequiredInputArgument {
	Type: string[];
	Name: string[];
}

export interface RequiredInputArguments {
	RequiredInputArgument: RequiredInputArgument[];
}

export interface Range {
	Minimum: string[];
	Maximum: string[];
	SuggestedStart: string[];
	SuggestedEnd: string[];
	SuggestedIncrement: string[];
}

export interface OptionalInputArgument {
	Name: string[];
	ShortDescription: string[];
	Type: string[];
	Range: Range[];
	DefaultValue: string[];
}

export interface OptionalInputArguments {
	OptionalInputArgument: OptionalInputArgument[];
}

export interface Flags {
	Flag: string[];
}

export interface OutputArgument {
	Type: string[];
	Name: string[];
	Flags: Flags[];
}

export interface OutputArguments {
	OutputArgument: OutputArgument[];
}

export interface FinancialFunction {
	Abbreviation: string[];
	CamelCaseName: string[];
	ShortDescription: string[];
	GroupId: string[];
	Flags: Flags[];
	RequiredInputArguments: RequiredInputArguments[];
	OptionalInputArguments: OptionalInputArguments[];
	OutputArguments: OutputArguments[];
}

export interface FinancialFunctions {
	FinancialFunction: FinancialFunction[];
}

export interface TaFuncApiXml {
	FinancialFunctions: FinancialFunctions;
}
