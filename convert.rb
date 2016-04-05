#!/usr/bin/ruby
# -*- coding: utf-8 -*-


require 'fileutils'
require 'rexml/document'


class Linebuffer
  MAXLINELEN=80
  def initialize
    @writebuffer=''
  end

  def append(line)
    @writebuffer+=line
  end

  def flush_buffer
    if @writebuffer.length>0
      @writebuffer+="\n"
      do_write
    end
  end

  def extract_head(n,from='')
    retVal=nil
    if n==0
      retVal="\n"
       @writebuffer[0,1]=''
    elsif n==1
      retVal=@writebuffer[0]+"\n"
      @writebuffer[0,2]=''
    else
      retVal=@writebuffer[0,n]+"\n"
      @writebuffer[0,n+1]=''
    end
    puts "extract_head(#{n},#{from}) >#{retVal}<  #=#{retVal.length}" if $DEBUG
    return retVal
  end
  
  def extract_line
    eol=@writebuffer.index("\n")
    if eol.nil? or eol>=MAXLINELEN
      return break_line
    end
    return extract_head(eol,'extract_line')
  end
  
  
  def break_line
    retVal=''
    return retVal if length<MAXLINELEN
    lastspace=@writebuffer.rindex(' ',MAXLINELEN)
    lastspace=@writebuffer.index('\n') if lastspace.nil?
   # lastspace=@writebuffer.index(' ')  if lastspace.nil?
    unless lastspace.nil?
      retVal=extract_head(lastspace,'break_line')
    end
    return retVal
  end
  
  def length
    return @writebuffer.length
  end

  def flush
    retVal=""
    l=length+1
    while l<length
      l=length
      retVal+=extractLine
    end
    retVal+=@writebuffer
    @writebuffer=''
    return retVal
  end
end

class Converter
  FIGURE_RE=/\(Abb\.\s*([0-9]{1,2})\)/
  CITATION_YEAR_RE=/\(([A-Za-z0-9\/"'`\s\\\-\{\}]+)\s+([0-9]{2,4}\s*[a-z]?)\s*:\s+([0-9a-zA-Z\.\s\-]+)\)/
  CITATION_NOYEAR_RE=/\(([A-Za-z0-9\/"'`\s\\\-\{\}]+)\s+([Oo]\.[Jj]\.\s*[a-z]?)\s*:\s*([0-9a-zA-Z\.\s\-]+)\)/

  SUFFIX_ADD=[
    {:pat=>Regexp.new("interview",Regexp::IGNORECASE),      :suffix=>':i'},
    {:pat=>Regexp.new("pers.+gespr.+",Regexp::IGNORECASE),  :suffix=>':c'},
  ]
  
  def initialize
    @doc=nil
    @f=nil
    @counter=0
    @italic_cache=nil
    @color_cache=nil
    @bold_cache=nil
    @indent_cache=nil
    @linebuffer=Linebuffer.new
    @appendix=[]
    @citationList=[]
    FileUtils.rm_rf './tmp'
    FileUtils.mkdir_p './tmp'
    FileUtils.chdir './tmp'
    system('unzip  ../odf.odt content.xml')
    File.open('content.xml','r:UTF-8') do |f|
      xml=f.read
      @doc = REXML::Document.new(xml)
    end
    @doc.each_element('//text:bookmark-start') do |t|
      t.parent.delete t
      puts t
    end
    @doc.each_element('//text:bookmark-end') do |t|
      t.parent.delete t
    end
    @doc.each_element('//text:soft-page-break') do |t|
      t.parent.delete t
    end
    
#    exit
  end

  def office_text
    return @doc.get_elements('/office:document-content/office:body/office:text')[0]
  end

  def italic_styles
    if @italic_cache.nil?
      @italic_cache=@doc.get_elements('/office:document-content/office:automatic-styles/*[@style:parent-style-name="Absatz-Standardschriftart"]/style:text-properties[@fo:font-style="italic"]').map { |x| x.parent.attribute('name').value}
    end
    return @italic_cache
  end

  def bold_styles
    if @bold_cache.nil?
      @bold_cache=@doc.get_elements('/office:document-content/office:automatic-styles/*[@style:parent-style-name="Absatz-Standardschriftart"]/style:text-properties[@fo:font-weight="bold"]').map { |x| x.parent.attribute('name').value}
    end
    return @bold_cache
  end

    def color_styles
    if @color_cache.nil?
      colors=@doc.get_elements('/office:document-content/office:automatic-styles/*[@style:parent-style-name="Absatz-Standardschriftart"]/style:text-properties[@fo:color]').map do |x|
        { :k=>x.parent.attribute('name').value, :v=>x.attribute('color').value }
      end
      @color_cache=Hash.new
      colors.each do |kv|
        @color_cache[kv[:k]]=kv[:v]
        puts "Style name: #{kv[:k]} color:#{kv[:v]}"
      end
    end
    return @color_cache
  end

  def indent_paragraph_styles
    if @indent_cache.nil?
      @indent_cache=@doc.get_elements('/office:document-content/office:automatic-styles/*[@style:family="paragraph"]/style:paragraph-properties[@fo:margin-left]').map { |x| x.parent.attribute('name').value}
    end
    return @indent_cache
  end


  
  def german_to_tex(line)
    return line.
            gsub(/ü/, '\"u').gsub(/Ü/,'\"U').
            gsub(/ä/, '\"a').gsub(/Ä/,'\"A').
            gsub(/ö/, '\"o').gsub(/Ö/,'\"O').
            gsub(/ß/, '\ss{}')
    
    
  end
  
  def quotes_to_tex(line)
    return line.sub(/‚/,"`").gsub(/’/,"'").gsub(/‘/,"'").gsub(/–/,'--').
            gsub(/„/,"``").gsub(/“/,"''").gsub(/”/,"''").
            gsub(/«/,'<<').gsub(/»/,'>>').gsub(/`` /,"`` \\todoconv{extra space?}").gsub(/ ''/,"\\todoconv{extra space?} ''")
  end

  def symbols_to_tex(line)
    return line.gsub(/%/,'\%').gsub(/&/,'\\\\&').gsub(/…/,'\\ldots').gsub(/([a-zA-Z])\\.\\.\\. /,'\1\\ldots ').gsub(/_/,'\_').gsub(/#/,'\#')
  end

  def remarks_for_tex(line)
    return line.gsub(/`` /,"``\\todoconv{extra space?}").gsub(/ ''/,"\\todoconv{extra space?}''").gsub(/,[A-Z]/,".\\todoconv{missing space?} ").gsub(/ \./,"\\todoconv{extra space?}.")
  end

  def imageref(line)
    puts "imageref(nil)" if line.nil?
    line=line.gsub(FIGURE_RE) do |img|
      key=$1
      puts "Figure #{img}"
      @appendix.push("\\input{./include/figure#{key}.tex}")
      core_ref="<,,,Abb.~\\ref{figure:#{key}}"
      "(#{core_ref})"
    end
    return line
  end

  def unpackimageref(line)
    return line.gsub(/<,,,Abb\./,"Abb.")
  end
  
  def citation(line)
    puts "citation(nil)" if line.nil?
    [CITATION_YEAR_RE,CITATION_NOYEAR_RE].each do |citation_re|
      line=line.gsub(citation_re) do |name|
        key=$1
        year=$2
        info=$3
        info.gsub!(/([0-9])-([0-9])/) { |x| $1+"--"+$2 } unless info.nil?
        year=year.gsub(/\s+/,"").gsub(/o\.J\./i,"0000")
        key=key.gsub(/\{\}/,'').gsub(/\s+/,'-').gsub(/[\"\\]/,'')+":#{year}"
        key.downcase!
        SUFFIX_ADD.each do |p|
          key=key+=p[:suffix]  unless p[:pat].match(info).nil?
        end
        if info.empty?
          r="~\\citep{#{key}}"
        else
          r="~\\citep[#{info}]{#{key}}"
        end
        @citationList.push("#{key}, #{name}") 
        r
      end
    end
    return line
  end
  
  def texify(line)
    line=line.gsub(/"/,"\\todoconv{check quotes} `'{} ")
    puts "line.nil?==true" if line.nil?
    line=german_to_tex(line)
    puts "line.nil?==true" if line.nil?
    line=quotes_to_tex(line)
    puts "line.nil?==true" if line.nil?
    line=symbols_to_tex(line)
    puts "line.nil?==true" if line.nil?
    line=remarks_for_tex(line)
    puts "line.nil?==true" if line.nil?
    line=imageref(line)
    line=citation(line)
    line=unpackimageref(line)
    return line
  end
  
  def newFile
    @counter+=1
    unless @f.nil?
      @appendix.each do |a|
        @f.write("#{a}\n")
      end
      @appendix=[]

      
      @f.write(@linebuffer.flush) 
      @f.close 
    end
    filename="section_#{sprintf("%02d",@counter)}.tex"
    open("sections.tex","a") do |t|
      t.puts("\n\\input{#{filename}}")
    end
    
    @f=File.open(filename,"w:UTF-8")
  end
  
  def do_write
    l=@linebuffer.extract_line
    while not l.empty?
      @f.write(l)
      l=@linebuffer.extract_line
    end
  end
  
  def write(l)
    newFile  if @f.nil?
    @linebuffer.append(l)
    do_write
  end

  def node_dispatch(elm)
    node_text(elm) if elm.is_a?(REXML::Text)
    if elm.is_a?(REXML::Element)
      node_anchor(elm) if elm.name=="h"
      node_span(elm) if elm.name=="span"
      node_p(elm) if elm.name=="p"
      node_s(elm) if elm.name=="s"
    end
  end


  def node_outline_anchor(elm)
    level=elm.attribute("outline-level").value.to_i
    title_node=elm.get_elements('text:span[string-length(text()) > 0]')
    puts "TITLE ELEMENTS #{title_node}  #=#{title_node.size}"
    title=''
    unless title_node.empty?
      title_node.first.each_child do |tit|
        if tit.kind_of?(REXML::Text)
          title+=tit.value
        else
          title+=" " if tit.name='s'
        end
        puts title
      end
    end
    puts "TITLE  #{title}"
    unless title.empty?
      if level==1
        newFile
        write("\\section{#{texify(title)}}\n")
      elsif level==2
        title.gsub!(/[0-9\.]+\s+/,'')
        write("\\subsection{#{texify(title)}}\n")
      elsif level==3
        title.gsub!(/[0-9\.]+\s+/,'')
        write("\\subsubsection{#{texify(title)}}\n")
      end
    end
  end
  
  def node_anchor(elm)
    puts "#h #{elm}"
    level=elm.attribute("outline-level")
    if  level.nil?
      elm.each_child do |c|
        node_dispatch(c)
      end
    else
      node_outline_anchor(elm)
    end
  end
  
  def node_span(elm)
    puts "#span #{elm}"
    style_name= elm.attribute("style-name")
    italic=italic_styles.include?(style_name.value)
    bold=bold_styles.include?(style_name.value)
    color=nil
    if color_styles.has_key?(style_name.value)
      color=color_styles[style_name.value]
    end
    color=nil # switch off
    write("{\\bfseries ") if bold
    write("\\emph{")  if italic
    write("{\\color{orange}") unless color.nil?
    elm.each_child do |c|
      node_dispatch(c)
    end
    write("}") unless color.nil?
    write("}")  if italic
    write("}")  if bold
  end

  def node_text(elm)
    write(texify(elm.value))
  end

  def node_p(elm)  # ends with empty line
    puts "#p #{elm}"
    style_name= elm.attribute("style-name")
    puts "STYLE-NAME #{style_name}"
    indentparagraph=indent_paragraph_styles.include?(style_name.value)
    write("\\paragraphIndent{")  if indentparagraph
    elm.each_child do |c|
      node_dispatch(c)
    end
    write("}")  if indentparagraph
    write("\n\n")
  end
  
  def node_s(elm) 
    write(' ');
  end
  
  def iterate
    office_text.each_child do |c|
      node_dispatch(c)
    end
    unless @f.nil?
      @f.write(@linebuffer.flush)
      @f.close 
    end
    File.open("citation_list.log","w") do |f|
       @citationList.each { |d| f.puts(d)}
    end

    
  end
  
end

a=Converter::new
puts "Italic "
puts a.italic_styles

puts "Bold "
puts a.bold_styles

puts "Color"
puts a.color_styles

puts "Indent"
puts a.indent_paragraph_styles

a.iterate


