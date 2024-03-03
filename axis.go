package xlsx

type Axis string

func (a Axis) C() (col, row int) {
	col, row = CellNameToCoordinates(string(a))
	return
}

func (a Axis) String() string {
	return string(a)
}

func NewAxis(col, row int) Axis {
	return Axis(CoordinatesToCellName(col, row))
}
