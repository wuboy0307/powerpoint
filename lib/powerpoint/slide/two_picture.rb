require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class TwoPicture
      include Powerpoint::Util

      attr_reader :title, :image_name, :image_path, :coords, :image_name2, :image_path2, :coords2

      def initialize(options={})
        require_arguments [:presentation, :title, :image_path, :image_path2], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @coords = default_coords
        @coords2 = default_coords2
        @image_name = File.basename(@image_path)
        @image_name2 = File.basename(@image_path2)
      end

      def save(extract_path, index)
        copy_media(extract_path, @image_path)
        copy_media(extract_path, @image_path2)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        File.extname(image_name).gsub('.', '')
      end

      def file_type2
        File.extname(image_name2).gsub('.', '')
      end

      def default_coords
        start_x = pixle_to_pt(0)
        default_width = pixle_to_pt(300)

        return {} unless dimensions = FastImage.size(image_path)
        image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}
        new_width = default_width < image_width ? default_width : image_width
        ratio = new_width / image_width.to_f
        new_height = (image_height.to_f * ratio).round
        {x: start_x, y: pixle_to_pt(120), cx: new_width, cy: new_height}
      end
      private :default_coords

      def default_coords2
        start_x = pixle_to_pt(360)
        default_width = pixle_to_pt(300)

        return {} unless dimensions = FastImage.size(image_path2)
        image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}
        new_width = default_width < image_width ? default_width : image_width
        ratio = new_width / image_width.to_f
        new_height = (image_height.to_f * ratio).round
        {x: start_x, y: pixle_to_pt(120), cx: new_width, cy: new_height}
      end
      private :default_coords2

      def save_rel_xml(extract_path, index)
        render_view('two_picture_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('two_picture_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end


end
