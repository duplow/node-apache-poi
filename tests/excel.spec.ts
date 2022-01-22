import { IExcel, Excel, IDocument } from '../index'

var excel : IExcel = new Excel()
excel.setFile('C:/Users/admin/Downloads/DescricaoLongaEN.xlsx')
excel.open()
  .then((doc) => {
    console.log({ doc, excel });
  })
  .catch(console.error)