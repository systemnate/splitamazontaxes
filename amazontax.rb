require 'csv'
require 'spreadsheet'

# Simple class to hold a CSV::Row and the order_id
class OrderEntity
  attr_accessor :order_id, :csv_row

  def initialize(order_id)
    @order_id = order_id
    @csv_row = []
  end
end

# Workhorse class to split the provided CSV file into a .xls
# file with each tab representing its own state
class SplitTaxes
  attr_reader :filename, :taxes, :dest_filename
  attr_accessor :orders, :current_order_entity, :states

  def initialize(filename, dest_filename)
    @filename = filename
    @dest_filename = dest_filename
    @taxes = CSV.read(@filename, headers: true)
    @orders = []
    @current_order_entity = nil
    @states = {}
    @state_lookup = { "New jersey" => "NJ", "South carolina" => "SC", "SOUTH CAROLINA" => "SC", "NEW JERSEY" => "NJ", "Alabama" => "AL", "Alaska" => "AK", "Alberta" => "AB", "American Samoa" => "AS", "Arizona" => "AZ", "Arkansas" => "AR", "Armed Forces (AE)" => "AE", "Armed Forces Americas" => "AA", "Armed Forces Pacific" => "AP", "British Columbia" => "BC", "California" => "CA", "Colorado" => "CO", "Connecticut" => "CT", "Delaware" => "DE", "District Of Columbia" => "DC", "Florida" => "FL", "Georgia" => "GA", "Guam" => "GU", "Hawaii" => "HI", "Idaho" => "ID", "Illinois" => "IL", "Indiana" => "IN", "Iowa" => "IA", "Kansas" => "KS", "Kentucky" => "KY", "Louisiana" => "LA", "Maine" => "ME", "Manitoba" => "MB", "Maryland" => "MD", "Massachusetts" => "MA", "Michigan" => "MI", "Minnesota" => "MN", "Mississippi" => "MS", "Missouri" => "MO", "Montana" => "MT", "Nebraska" => "NE", "Nevada" => "NV", "New Brunswick" => "NB", "New Hampshire" => "NH", "New Jersey" => "NJ", "N.j." => "NJ", "New Mexico" => "NM", "New York" => "NY", "Newfoundland" => "NF", "North Carolina" => "NC", "North Dakota" => "ND", "Northwest Territories" => "NT", "Nova Scotia" => "NS", "Nunavut" => "NU", "Ohio" => "OH", "Oklahoma" => "OK", "Ontario" => "ON", "Oregon" => "OR", "Pennsylvania" => "PA", "Prince Edward Island" => "PE", "Puerto Rico" => "PR", "Quebec" => "PQ", "Rhode Island" => "RI", "Saskatchewan" => "SK", "South Carolina" => "SC", "South Dakota" => "SD", "Tennessee" => "TN", "Texas" => "TX", "Utah" => "UT", "Vermont" => "VT", "Virgin Islands" => "VI", "Virginia" => "VA", "Washington" => "WA", "West Virginia" => "WV", "Wisconsin" => "WI", "Wyoming" => "WY", "Yukon Territory" => "YT" } 
  end

  def run
    puts "Running..."
    populate_orders
    break_orders_into_states
    flatten_orders
    write_to_xls
  end


  def populate_orders
    puts "Populating Orders..."
    @taxes.each do |tax|
      if !(tax['Order_ID'].nil?)
        @current_order_entity = OrderEntity.new(tax['Order_ID'])
        @current_order_entity.csv_row << tax
        @orders << current_order_entity
      else
        @current_order_entity.csv_row << tax
      end      
    end
  end

  def break_orders_into_states
    puts "Breaking orders into states..."
    @orders.each do |order|
      order.csv_row.each do |row|
        ship_to_state = @state_lookup[row['Ship_To_State'].to_s.capitalize] ||= row['Ship_To_State']
        if ship_to_state
          if @states.has_key?(ship_to_state)
            @states[ship_to_state] << order.csv_row
          else
            @states[ship_to_state] = []
            @states[ship_to_state] << order.csv_row
          end
        end
      end
    end
  end

  def flatten_orders
    puts "Flattening orders..."
    @states.each do |k,v|
      @states[k] = @states[k].flatten
    end
  end

  def write_to_xls
    puts "Writing to XLS..."
    book = Spreadsheet::Workbook.new
    @states.each do |state, value|
      sheet = book.create_worksheet :name => state
      @states[state].each_with_index do |taxes,i|
        taxes.headers.each do |header|
          sheet.row(i).push header
        end
        taxes.to_a.each do |tax|
          sheet.row(i+1).push tax[1]
        end
      end
    end
    book.write @dest_filename
  end
end

SplitTaxes.new("../2015Combined.csv", "../Combined.xls").run