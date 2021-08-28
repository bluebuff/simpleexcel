package simpleexcel

type Handler func(styleMng StyleManager, cell *Cell)

func CompareLessAndNumberConditionInt32Style(compareValue int32) Handler {
	return func(styleMng StyleManager, cell *Cell) {
		if v, ok := cell.Value.(int32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndNumberConditionInt64Style(compareValue int64) Handler {
	return func(styleMng StyleManager, cell *Cell) {
		if v, ok := cell.Value.(int64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndNumberConditionUint32Style(compareValue uint32) Handler {
	return func(styleMng StyleManager, cell *Cell) {
		if v, ok := cell.Value.(uint32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndNumberConditionUint64Style(compareValue uint64) Handler {
	return func(styleMng StyleManager, cell *Cell) {
		if v, ok := cell.Value.(uint64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndDecimalsConditionFloat32Style(compareValue float32) Handler {
	return func(styleMng StyleManager, cell *Cell) {
		if v, ok := cell.Value.(float32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(DecimalsCondition)
		}
	}
}

func CompareLessAndDecimalsConditionFloat64Style(compareValue float64) Handler {
	return func(styleMng StyleManager, cell *Cell) {
		if v, ok := cell.Value.(float64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(DecimalsCondition)
		}
	}
}
