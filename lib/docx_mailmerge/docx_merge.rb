module DocxMailmerge
  class DocxMerge
    MISSING_VALUE_TEXT = "***MISSING VALUE***"
    attr_reader :doc, :data

    def initialize(data)
      @data = data
    end

    def field_names
      (simple_field_names + complex_field_names).uniq
    end

    def merge(file)
      @doc = Nokogiri::XML(file)
      simple_merge
      complex_merge
      @doc.to_xml
    end

    private

    def simple_merge_nodes
      @doc.xpath("//w:fldSimple[contains(@w:instr,'MERGEFIELD')]")
    end

    def complex_merge_nodes
      @doc.xpath("//w:instrText[contains(text(),'MERGEFIELD')]")
    end

    def simple_field_names
      simple_merge_nodes.map do |simple_node|
        first_mergefield_name simple_node["w:instr"]
      end
    end

    def complex_field_names
      complex_merge_nodes do |complex_node|
        first_mergefield_name complex_node.content
      end
    end

    def simple_merge
      simple_merge_nodes.each do |simple_node|
        ft = field_text(simple_node["w:instr"])
        simple_node.search(".//w:t").first.content = ft
        simple_node.replace(simple_node.children)
      end
    end

    def complex_merge
      complex_merge_nodes.each do |complex_node|
        # begin tag
        complex_node.parent.previous_element.remove

        # separator tag
        complex_node.parent.next_element.remove

        text_node = complex_node.parent.next_element
        text_node.search(".//w:t").first.content = field_text(complex_node.content)

        # end tag and potientally more extra junk
        search_result = ""
        while text_node.next_element && (search_result.nil?  || search_result.empty?)
          search_result = text_node.next_element.search('.//w:fldChar[@w:fldCharType="end"]')
          text_node.next_element.remove
        end

        # mergfield tag
        complex_node.parent.remove
      end
    end

    def field_text(node)
      field_name = first_mergefield_name(node)
      @data[field_name] || "#{MISSING_VALUE_TEXT}#{field_name}"
    end

    def first_mergefield_name(node)
      node.match(/ MERGEFIELD \"(.*)\"/)[1]
    end

  end
end