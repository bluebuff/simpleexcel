package handle

import (
	"github.com/bluebuff/simple-excelize/core/common"
	"github.com/bluebuff/simple-excelize/core/context"
)

// 统一设置指定的列宽（在布局范围内）
func SetLayoutColWidth(width float64) context.Handler {
	if width == 0 {
		width = common.DEFAULT_COLUMN_WIDTH
	}
	return func(ctx context.Context) error {
		layout := ctx.GetLayout()
		ctx.SetColWidth(layout.Left, layout.Right-1, width)
		return nil
	}
}
