class DocxDownloadService
  # include ActiveStorage::Downloading
  attr_reader :blob

  def initialize(blob)
    @blob = blob
  end

  def doc
    download_blob_to_tempfile do |file|
      Docx::Document.open(file)
    end
  end
end