const excelToJson = require('convert-excel-to-json');

const main = async () => {
    let count = 1
    let arr = []
    let thisRow = []

    const result = excelToJson({
        sourceFile: 'tips1-4.xlsx'
    })

    // console.log(result.Hoja1)

    result.Hoja1.forEach((el, i)=>{
        if (el.A === `${count}.`) {
            if (thisRow[0]) {
                arr.push(thisRow)
            }

            count++
            thisRow = []
            thisRow.push(el)

            // thisRow = []
            // count++

            // thisRow.push(el)
        } else if (i !== 0 && i !== 1) {
            thisRow.push(el)
        }

        // thisRow.push(el)
    })

    // console.log(arr[0])
    // console.log(arr[1])
    // console.log(arr[2])
    // console.log(arr[3])

    let res = parseVector(arr)

    console.log(res)
}

const parseVector = (vector) => {
    let arr = vector.reduce((acc, el, i)=> {
        let thisRow = el.reduce((accRow, elRow, elI)=> {
            if (elRow.A) {
                accRow.NUMBER = elRow.A.replace('.', '')
            }

            if (elRow.B) {
                accRow.TIP = elRow.B
            }

            if (elRow.C) {
                if (!accRow.QUESTION_AL) {
                    accRow.QUESTION_AL = elRow.C
                } else {
                    accRow.QUESTION_VF = elRow.C
                }
            }

            if (elRow.D) {
                if (!accRow.VARIABLES) accRow.VARIABLES = []

                accRow.VARIABLES.push(elRow.D)
            }

            if (elRow.E && elRow.E !== '') {
                accRow.VARIABLES_SOLUTION = elRow.E
            }

            if (elRow.G && elRow.G !== '') {
                accRow.VF_SOLUTION = elRow.G
            }

            return accRow
        }, {})

        acc.push(thisRow)

        return acc
    }, [])

    return arr
}

main()