package common

type ConditionExpress string

const (
	NotConditionExpression      ConditionExpress = ""
	NumberConditionExpression   ConditionExpress = `[{"type":"cell","criteria":"<","format":%d,"value":"%.2f"}]`
	DecimalsConditionExpression ConditionExpress = `[{"type":"cell","criteria":"<","format":%d,"value":"%.2f"}]`
)
