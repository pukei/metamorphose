require 'rubyXL'

module Metamorphose

  def execute(args = {})
    workbook = RubyXL::Parser.parse('o.xlsx')
    worksheet = workbook.sheet('Sheet1')
    # worksheet = workbook.sheet(0)
    row = 16
    column = 2 # C

    value = 'BlahBlahBlah'
    worksheet.add_cell(row, column, value)
    workbook.save

    cell = worksheet[row][column]
    cell.raw_value
    cell.value
  end

end
