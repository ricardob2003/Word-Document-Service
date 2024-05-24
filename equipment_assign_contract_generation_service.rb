# app/services/contract_document_generator.rb
class EquipmentAssignContractGenerationService
  def initialize(contract, employee_name, employee, active_contract_equipment)
    @contract = contract
    @employee_name = employee_name
    @employee = employee
    @active_contract_equipment = active_contract_equipment
  end

  def generate
    template_path = Rails.root.join("public", "contract_templates", "Contrato_de_Prestamo_de_Equipo.docx").to_s
    doc = Docx::Document.open(template_path)

    # Log equipment details if present
    if @active_contract_equipment.present?
      @active_contract_equipment.each do |equipment|
        puts "Equipo: #{equipment.equipment_name}, Modelo: #{equipment.equipment_model}, Código: #{equipment.equipment_code}, Costo: #{equipment.equipment_cost}, Asignado a: #{equipment.employee_name} #{equipment.employee_apellido}"
      end
    else
      puts "No se encontraron equipos asignados al empleado."
    end

    # Assuming Contract.replace_template_placeholders is a class method you've defined elsewhere
    Contract.replace_template_placeholders(doc, @employee, @contract)

    fill_table_with_equipment(doc.tables[0]) if doc.tables[0].rows.size >= 1
    substitute_employee_details(doc.tables[1])

    save_and_attach_document(doc)
  end

  private

  def fill_table_with_equipment(table)
    reference_row = table.rows[0]

    @active_contract_equipment.each_with_index do |equipment, index|
      new_row = reference_row.copy
      new_row.insert_after(reference_row)

      fill_equipment_row(new_row, equipment)
    end
  end

  def fill_equipment_row(row, equipment)
    row.cells.each do |cell|
      cell.paragraphs.each do |paragraph|
        paragraph.each_text_run do |text|
          substitute_equipment_text(text, equipment)
        end
      end
    end
  end

  def substitute_equipment_text(text, equipment)
    case text.to_s
    when "Codigo"
      text.substitute("Codigo", equipment.equipment_code)
    when "Modelo"
      text.substitute("Modelo", equipment.equipment_model)
    when "Marca"
      text.substitute("Marca", equipment.equipment_name.capitalize)
    when "Precio"
      text.substitute("Precio", "#{equipment.equipment_cost} Lps.")
    end
  end

  def substitute_employee_details(table)
    cell_to_replace = table.rows[1].cells[1]
    cell_to_replace.paragraphs.each do |paragraph|
      paragraph.each_text_run do |text|
        text.substitute("employee_rtn2", @employee.rtn.to_s)
      end
    end
  end

  def save_and_attach_document(doc)
    tmp_file_path = Rails.root.join("tmp", "contract_#{@contract.id}_#{@employee_name}_#{Date.today}.docx")
    doc.save(tmp_file_path)

    @contract.document.attach(io: File.open(tmp_file_path), filename: "Contrato de equipo de #{@employee_name}.docx")
    @contract.save

    File.delete(tmp_file_path) if File.exist?(tmp_file_path)
  end
end
