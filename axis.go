package xlsx

type Axis string

func (a Axis) C() (col, row int) {
	col, row = CellNameToCoordinates(string(a))
	return
}
