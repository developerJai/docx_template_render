class DocumentsController < ApplicationController

  before_action :set_document, only: %i[ show edit update destroy doc_preview]

  # GET /documents or /documents.json
  def index
    @documents = Document.all
  end

  # GET /documents/1 or /documents/1.json
  def show
    respond_to do |format|
      format.html
      format.docx do

        @file_temp = Tempfile.new("document_tmp_#{@document.id}", "#{Rails.root}/tmp")
        @file_temp.binmode
        
        @document.doc.download {|chunk| @file_temp.write(chunk) }

        # Initialize DocxReplace with your template
        doc = DocxReplace::Doc.new(@file_temp.path, "#{Rails.root}/tmp")

        # Replace some variables. $var$ convention is used here, but not required.
        doc.replace("%{company_cin_no}", "This is cin number")
        doc.replace("%{company_address}", "This is the company address")
        doc.replace("%{company_contact_email}", "This is company email")

        # Replace multiple occurrences
        doc.replace("%{branch_name}", "This is branch name", true)

        # Write the document back to a temporary file
        # tmp_file = Tempfile.new("document_#{@document.id}", "#{Rails.root}/tmp")
        doc.commit#(tmp_file.path)
        # doc_pdf_dir = "#{Rails.root}/tmp/pdf_temp_#{@document.id}.pdf"
        # Libreconv.convert(@file_temp.path, doc_pdf_dir)

        # Respond to the request by sending the temp file
        send_file @file_temp.path, filename: "doc_#{@document.id}_report.docx", disposition: 'preview'
      end
    end
  end

  def doc_preview
    @file_temp = Tempfile.new("document_tmp_#{@document.id}", "#{Rails.root}/tmp")
    @file_temp.binmode    
    @document.doc.download {|chunk| @file_temp.write(chunk) }

    # Initialize DocxReplace with your template
    temp_doc = DocxReplace::Doc.new(@file_temp.path, "#{Rails.root}/tmp")

    # Replace some variables. $var$ convention is used here, but not required.
    temp_doc.replace("%{company_cin_no}", "This is cin number")
    temp_doc.replace("%{company_address}", "This is the company address")
    temp_doc.replace("%{company_contact_email}", "This is company email")

    # Replace multiple occurźences
    temp_doc.replace("%{branch_name}", "This is branch name", true)

    temp_doc.commit(@file_temp.path)

    # d = Docx::Document.open(@file_temp.path)


    # @data = d.to_html.html_safe
    @data = temp_doc.to_html

    File.delete(@file_temp.path) if File.exist?(@file_temp.path)

  end

  # GET /documents/new
  def new
    @document = Document.new
  end

  # GET /documents/1/edit
  def edit
  end

  # POST /documents or /documents.json
  def create
    @document = Document.new(document_params)

    respond_to do |format|
      if @document.save
        format.html { redirect_to document_url(@document), notice: "Document was successfully created." }
        format.json { render :show, status: :created, location: @document }
      else
        format.html { render :new, status: :unprocessable_entity }
        format.json { render json: @document.errors, status: :unprocessable_entity }
      end
    end
  end

  # PATCH/PUT /documents/1 or /documents/1.json
  def update
    respond_to do |format|
      if @document.update(document_params)
        format.html { redirect_to document_url(@document), notice: "Document was successfully updated." }
        format.json { render :show, status: :ok, location: @document }
      else
        format.html { render :edit, status: :unprocessable_entity }
        format.json { render json: @document.errors, status: :unprocessable_entity }
      end
    end
  end

  # DELETE /documents/1 or /documents/1.json
  def destroy
    @document.destroy

    respond_to do |format|
      format.html { redirect_to documents_url, notice: "Document was successfully destroyed." }
      format.json { head :no_content }
    end
  end

  private
    # Use callbacks to share common setup or constraints between actions.
    def set_document
      @document = Document.find(params[:id])
    end

    # Only allow a list of trusted parameters through.
    def document_params
      params.require(:document).permit(:doc)
      # params.fetch(:document, {})
    end
end
