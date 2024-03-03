package xlsx

type Axis string

func (a Axis) C() (col, row int) {
	col, row = CellNameToCoordinates(string(a))
	return
}

func (a Axis) String() string {
	return string(a)
}

func (a Axis) Col() int {
	col, _ := a.C()
	return col
}

func (a Axis) Row() int {
	_, row := a.C()
	return row
}

func (a Axis) Add(col, row int) Axis {
	c, r := a.C()
	return NewAxis(c+col, r+row)
}

func NewAxis(col, row int) Axis {
	return Axis(CoordinatesToCellName(col, row))
}
