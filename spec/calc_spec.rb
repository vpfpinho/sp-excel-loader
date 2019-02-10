
require 'sp-excel-loader'

RSpec.describe Sp::Excel::Loader do

  we = nil

  def clean_files
      File.delete('./spec/json/mcalc_spec.rb') if File.file?('./spec/json/mcalc_spec.rb')
      File.delete('./spec/json/model.json') if File.file?('./spec/json/model.json')
      File.delete('./spec/json/model_typed.json') if File.file?('./spec/json/model_typed.json')
      File.delete('./spec/json/tables/payroll_grant_types.json') if File.file?('./spec/json/tables/payroll_grant_types.json')
      Dir.rmdir('./spec/json/tables') if File.exists?('./spec/json/tables')
      Dir.rmdir('./spec/json') if File.exists?('./spec/json')
  end

  it 'create exporter instance' do
    clean_files()
    we = Sp::Excel::Loader::PayrollExporter.new('./spec/model.xls', false)
    we.export './spec/json/'
    expect(we).to be_an_instance_of(Sp::Excel::Loader::PayrollExporter)
  end

  context 'check model json files' do

    model = nil;

    it 'model.json was created' do
        expect(File.file?('./spec/json/model.json')).to eq(true)
        json  = File.read('./spec/json/model.json')
        model = JSON.parse(json)
    end

    it 'model has correct structure' do
      expect(model['values']).to be_an_instance_of(Hash)
      expect(model['formulas']).to be_an_instance_of(Hash)
      expect(model['lines']['header']).to be_an_instance_of(Hash)
      expect(model['lines']['formulas']).to be_an_instance_of(Array)
      expect(model['lines']['values']).to be_an_instance_of(Array)
    end

    it 'model has expected information' do
      expect(model['values']['I8']).to                      eq('COD_FUNC=12')
      expect(model['formulas']['I21']).to                   eq('DIAS_ANO_CALENDARIO=(DT_FIM_ANO)-(DATE(ANO_PROC,1,1))+1')
      expect(model['lines']['header']['I27']).to            eq('UNIT_VALUE')
      expect(model['lines']['header'].size).to              eq(16)
      expect(model['lines']['formulas'].size).to            eq(6)
      expect(model['lines']['values'].size).to              eq(6)
      expect(model['lines']['formulas'][0]['DESCRICAO']).to eq('D28=2+2')
      expect(model['lines']['values'][5]['COD']).to         eq('COD_SUBS_FERIAS_50="A006"')
    end

  end

  context 'check table json files' do

    table = nil;

    it 'table json was created' do
      expect(File.file?('./spec/json/tables/payroll_grant_types.json')).to eq(true)
      json  = File.read('./spec/json/tables/payroll_grant_types.json')
      table = JSON.parse(json)
    end

    it 'table has correct structure' do
      expect(table).to                 be_an_instance_of(Array)
      expect(table[0]).to              be_an_instance_of(Hash)
      expect(table[0]['data']).to      be_an_instance_of(Array)
      expect(table.size).to            eq(7)
      expect(table[0]['data'].size).to eq(12)
      expect(table[6]['data'].size).to eq(12)
    end

    it 'table has expected data' do
      expect(table[0]['name']).to     eq('id')
      expect(table[0]['type']).to     eq('text')
      expect(table[0]['data'][0]).to  eq('P_LIM_SREF_NUM')
      expect(table[0]['data'][11]).to eq('LIM_AC_I_MGESP')
      expect(table[6]['name']).to     eq('exemption_limit_social_security')
      expect(table[6]['type']).to     eq('number')
      expect(table[6]['data'][0]).to  eq(4.27)
      expect(table[6]['data'][11]).to eq(100.24)
      clean_files()
    end

  end

end
