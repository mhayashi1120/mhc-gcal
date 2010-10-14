#!/usr/bin/ruby
# -*- ruby -*-

## mhc-gcal
## Copyright (C) 2007 Masahiro Hayashi <mhayashi1120@gmail.com>
## Author: Masahiro Hayashi <mhayashi1120@gmail.com>

## "mhc2gcal" is base code of mhc-gcal
## Original author: Nao Kawanishi <river2470@gmail.com>

## "mhc2ical" is base code of mhc2gcal
## Original author: Yojiro UO <yuo@iijlab.net>
## "today" is base code of mhc2ical
## Original author: Yoshinari Nomura <nom@quickhack.net>

## "ol2gcal" is also base code of mhc2gcal
## Original author: <zoriorz@gmail.com>

#### Usage:
## You can use ~/.mhc-gcal file as default settings.
##
## Almost case, like following example works fine.
## For more details, see `set_config_value' method ;-)
##
## ---------- Sample ------------
## gcal_mail = mhayashi1120@gmail.com
## gcal_pass = password
## ---------- Sample End --------

#### TODO:
# MHC-Alarm
# delete if mhc intersect item


require 'rubygems'
require 'googlecalendar/calendar'
require 'date'
require 'mhc-schedule'
require 'mhc-date'
require 'nkf'

# todo error check input date value
class MhcGcalConfig

  class ParseError < StandardError; end #:nodoc: all

  DEF_RCFILE = "~/.mhc-gcal"

  attr_reader :GCAL_FEED, :GCAL_MAIL, :GCAL_PASS
  attr_reader :date_from
  attr_reader :date_to
  attr_reader :category
  attr_reader :secret_categories
  attr_reader :time_from
  attr_reader :time_to
  attr_reader :title_display_where
  attr_reader :secret_title
  attr_reader :mhc_schedule_files
  attr_reader :mhc_schedule_dir
  attr_reader :http_proxy_server
  attr_reader :http_proxy_port
  attr_reader :http_proxy_user
  attr_reader :http_proxy_pass

  def initialize(argv)

    @rcfile = nil
    set_rcfile(argv)

    @GCAL_FEED = nil
    @GCAL_MAIL = nil
    @GCAL_PASS = nil
    @date_from = @date_to = MhcDate .new
    @category  = '!Holiday'
    @secret_categories   = ["Private"]
    @title_display_where = false
    @secret_title = 'SECRET'
    @mhc_schedule_files = [ MhcScheduleDB::DEF_RCFILE ]
    @mhc_schedule_dir = MhcScheduleDB::DEF_BASEDIR

    read_rcfile()

    while option = argv .shift
      case (option)
      when /^--category=(.+)/
        @category = $1
      when /^--secret=(.*)/
        secret = $1
      when /^--date=([^+]+)(\+(-?[\d]+))?/
        @date_from, @date_to = string_to_date($1, $3) || usage()
      when /^--date=([^-]+)\-(.+)/
        @date_from, @date_to = string_to_date2($1, $2) || usage()
      else
        usage()
      end
    end

    @secret_categories = parse_category(secret) if secret
    @time_from = Time.mktime(@date_from.y.to_i, @date_from.m.to_i, @date_from.d.to_i, 0, 0, 0)
    @time_to = Time.mktime(@date_to.y.to_i, @date_to.m.to_i, @date_to.d.to_i, 23, 59, 59)

    usage() if @GCAL_FEED.nil?
    usage() if @GCAL_MAIL.nil?
    usage() if @GCAL_PASS.nil?

    dump_value() if $DEBUG

  end

  def secret_category?(cate)
    self .secret_categories .each do |c|
      regexp = Regexp .new(c, nil, "e")
      return true if regexp =~ cate.downcase
    end

    return false

  end

  ###################
  private
  ###################

  def set_rcfile(argv)

    argv.each do |arg|
      case (arg)
      when /^--rcfile=(.+)/
        @rcfile = $1
      end
    end

    @rcfile = File.expand_path(DEF_RCFILE) if @rcfile.nil?

  end

  def read_rcfile()

    unless FileTest.exist?(@rcfile)
      return
    end

    open(@rcfile, "r")  do |f|
      for line in f.readlines
        next if line =~ /^( *#| *$)/
        next unless line =~ /^ *([a-z_]+) *= *(.*)/i

        # todo coding-system

        set_config_value($1, $2)
      end
    end

  end

  def set_config_value(directive, value)
    case directive.downcase
    when "gcal_feed"
      @GCAL_FEED = value
    when "gcal_mail"
      @GCAL_MAIL = value
      if @GCAL_FEED.nil?
        @GCAL_FEED = "http://www.google.com/calendar/feeds/#{value}/private/full" 
      end
    when "gcal_pass"
      @GCAL_PASS = value
    when "http_proxy"
      raise(ParseError, "http_proxy name is invalid.") unless value =~ /^([^:]+):([0-9]+)$/
      @http_proxy_server = $1
      @http_proxy_port = $2.to_i
    when "http_proxy_user"
      @http_proxy_user = value
    when "http_proxy_pass"
      @http_proxy_pass = value
    when "date_from"
      @date_from = parse_config_date(value)
    when "date_to"
      @date_to = parse_config_date(value)
    when "category"
      #todo
      @category = parse_category(value)
    when "secret_categories"
      @secret_categories = parse_category(value)
    when "title_display_where"
      case value 
      when "true"
        @title_display_where = true
      when "false"
        @title_display_where = false
      else
        raise(ParseError, "Title_Display_Where must be `true' or `false' ")
      end
    when "secret_title"
      @secret_title = value
    when "mhc_schedule_files"
      @mhc_schedule_files = parse_file_list(value)
    when "mhc_schedule_dir"
      @mhc_schedule_dir = parse_file(value)
    else
      raise(ParseError, "Unrecognized directive #{directive} found.")
    end
  end

  def parse_config_date(value)

    diff_date = 0
    today = Date.today

    case value
    when /^today([+-][0-9]+)?$/
      diff_date = $1.to_i
    when /^thismonth([+-][0-9]+)?$/
      succ_month = (Date.new(today.year, today.month) >> $1.to_i)
      diff_date = (succ_month - today).to_i
    when /^thisyear([+-][0-9]+)?$/
      succ_year = (Date.new(today.year) >> ($1.to_i * 12))
      diff_date = (succ_year - today).to_i
    else
      raise(ParseError, "Invalid date value.")
    end

    return MhcDate .new(today.strftime("%Y%m%d")) .succ(diff_date)

  end

  def parse_category(value)
    return value .split .collect { |x| x .downcase }
  end

  def parse_file_list(value)
    ret = []
    value .split .each do |file|
      ret << parse_file(file)
    end

    return ret

  end

  def parse_file(value)
    return File.expand_path(value)
  end

  def string_to_date(base, range)
    ret_from = nil
    ret_to   = nil

    case base .downcase
    when 'today'
      ret_from = MhcDate .new
    when 'tomorrow'
      ret_from = MhcDate .new .succ
    when /^(sun|mon|tue|wed|thu|fri|sat)/
      ret_from = MhcDate .new .w_this(base .downcase)
    when /^\d{8}$/
      ret_from = MhcDate .new(base)
    when /^\d{6}$/
      ret_from = MhcDate .new(base + '01')
      if range
        ret_to = ret_from .succ(range .to_i)
      else
        ret_to = MhcDate .new(base + format("%02d", ret_from .m_days))
      end
    when /^\d{4}$/
      ret_from = MhcDate .new(base + '0101')
      if range
        ret_to = ret_from .succ(range .to_i)
      else
        ret_to = MhcDate .new(base + '1231')
      end
    else
      return nil
    end

    ret_to   = ret_from .succ((range || '0') .to_i) if ! ret_to

    return [ret_from, ret_to]

  end

  def string_to_date2(s1, s2)
    item = []
    [s1, s2] .each do |str|
      case str .downcase
      when 'today'
        item << MhcDate .new
      when 'tomorrow'
        item << MhcDate .new .succ
      when /^(sun|mon|tue|wed|thu|fri|sat)/
        item << MhcDate .new .w_this(str .downcase)
      when /^\d{8}$/
        item << MhcDate .new(str)
      when /^\d{6}$/
        item << MhcDate .new(str + '01')
      when /^\d{4}$/
        item << MhcDate .new(str + '0101')
      else
        item << nil
      end
    end

    return item

  end

  def dump_value()
    dump_item :GCAL_FEED
    dump_item :GCAL_MAIL
    dump_item :GCAL_PASS
    dump_item :date_from
    dump_item :date_to
    dump_item :category
    dump_item :secret_categories
    dump_item :time_from
    dump_item :time_to
    dump_item :title_display_where
    dump_item :secret_title
    dump_item :mhc_schedule_files
    dump_item :mhc_schedule_dir
    dump_item :http_proxy_server
    dump_item :http_proxy_port
    dump_item :http_proxy_user
    dump_item :http_proxy_pass
  end

  def dump_item(symbol)
    value = self.send(symbol)
    if value.kind_of?(Array)
      value2 = "[" + value.join(", ") + "]"
    else
      value2 = value.to_s
    end
    print(symbol.to_s + " = \"" + value2 + "\"\n")
  end

  def usage()
    STDERR .print <<EOF
usage: #{$0} [options]
  Upload your MHC schedule to Google Calendar:
  --help               show this message.
  --category=CATEGORY  pick only in CATEGORY. 
                       '!' and space separated multiple values are allowed.
  --secret=CATEGORY    change the title of the event to 'SECRET'
                       space separated multiple values are allowed.
  --date={string[+[-]n],string-string}
                       set a period of date.
                       string is one of these:
                         today, tomorrow, sun ... sat, yyyymmdd, yyyymm, yyyy
                       yyyymm lists all days in the month and yyyy lists all
                       days in the year.
                       list n+1 days of schedules if +n is given.
                       default value is 'today+0'
EOF
    exit 1
  end

end

def make_gtime(dt, tm = nil)

  if tm.nil?
    return Time.mktime(dt.y, dt.m, dt.d, 0, 0, 0)
  elsif tm.kind_of?(MhcTime)

    # MhcDate object of consider time
    cdt = dt.succ(tm.hh / 24)

    return Time.mktime(cdt.y, cdt.m, cdt.d, tm.hh % 24, tm.mm, 0)
  elsif tm.kind_of?(String) and tm =~ /[0-9]{1,2}:[0-5][0-9]/
    hh, mm = tm.split(":") .collect do |val| val.to_i end

    # MhcDate object of consider time
    cdt = dt.succ(hh / 24)

    return Time.mktime(cdt.y, cdt.m, cdt.d, hh % 24, mm, 0)
  else
    raise(TypeError, "argument `tm' is invalid.")
  end

end

# GoogleCalendar::Calendar wrapper
class GCalDb

  @events = nil
  @gcal = nil

  def initialize(conf)

    GoogleCalendar::Service::proxy_addr= conf.http_proxy_server
    GoogleCalendar::Service::proxy_port= conf.http_proxy_port
    GoogleCalendar::Service::proxy_user= conf.http_proxy_user
    GoogleCalendar::Service::proxy_pass= conf.http_proxy_pass

    srv = GoogleCalendar::Service.new(conf.GCAL_MAIL, conf.GCAL_PASS)
    @gcal = GoogleCalendar::Calendar.new(srv, conf.GCAL_FEED)

    @events = @gcal.events(:'start-min' => conf.time_from,
                           :'start-max' => conf.time_to,
                           :'max-results' => 100)

  end

  #todo slip off time
  def newer?(gev, log)
    return log.mtime > gev.updated
  end

  def find_upload_event(mev, lastlog)
    gev = self .find_gevent(mev.rec_id)

    return @gcal .create_event unless gev

    ## when entry was edited long time ago, lastlog is nil
    ## this case probably already uploaded.
    if lastlog && newer?(gev, lastlog)
      return gev
    else
      return nil
    end

  end

  def find_gevent(id)

    @events .each do | gev |
      if gev.extended_property["MHC-Record-Id"]
        if gev.extended_property["MHC-Record-Id"] == id
          return gev
        end
      end
    end

    return nil

  end

  def added_entries()

    ret = []

    @events .each do | gev |
      ret << gev if gev.extended_property["MHC-Record-Id"].nil?
    end

    return ret

  end

end

# todo add_sch when subject include newline, why this happen?

# MhcScheduleDB wrapper
class MhcDb

  @db = nil
  @searched = nil
  @log = nil
  @limited_entries = nil

  # db => MhcScheduleDB, conf => MhcGcalConfig
  def initialize(conf)
    @db = MhcScheduleDB .new(conf.mhc_schedule_dir, *conf.mhc_schedule_files)
    @searched = @db .search(conf.date_from, conf.date_to, conf.category)
    # NOTE: DO NOT USE `path' property because i use Mail folder local and via samba
    @log = MhcLog .new(File.expand_path(".mhc-db-log", conf.mhc_schedule_dir))

    # todo 5 month ago
    limit = (Date.today << 5).strftime("%Y%m%d")

    @limited_entries = []
    @log .entries .each do |ent|
      next if ent.mtime.strftime("%Y%m%d") < limit
      @limited_entries << ent 
    end

  end

  # todo
  def each_events()
    @searched .each do |date, mevs|
      mevs .each do |mev|
        #todo probablly ruby 1.8 or later
        lastlog = log_entries(mev).last
        yield(date, mev, lastlog)
      end
    end
  end

  # mev => MhcScheduleItem
  # return Array of MhcLogEntry
  def log_entries(mev)

    ret = []

    @limited_entries .each do |ent|
      next unless ent .rec_id == mev .rec_id
      ret << ent
    end

    return ret

  end

#   def find_mevent(id)
#     self.each_events do |mev|
#       return mev if mev.rec_id == id
#     end
#   end

  # return MhcLogEntry that recently deleted.
  def deleted_entries()
    ret = []
    @limited_entries.reverse.each do |ent|
      next if ent.status != "D"
      ret << ent 
    end

    return ret

  end

  # mev => MhcScheduleItem
  def add_sch(mev)
    return @db .add_sch(mev)
  end

end

def mevents_to_gevents(gev, mev, date)

  is_secret = @config.secret_category?(mev.category_as_string)
  if is_secret then
    gev.title = @config.secret_title
  else
    gev.title = NKF .nkf("-w", mev.subject)
  end

  if mev .location and mev .location != ''
    gev.where =  NKF .nkf("-w", mev.location)
    if ! is_secret && @config .title_display_where
      gev.title +=  '[' + gev.where + ']'
    end
  else
    gev.where =  ''
  end

  if mev.time_b
    gev.st = make_gtime(date, mev.time_b)
    if ! mev.time_e .nil?
      gev.en = make_gtime(date, mev.time_e)
    else
      gev.en = make_gtime(date, mev.time_b)
    end
  else
    gev.st = make_gtime(date)
    gev.en = make_gtime(date.succ)
    gev.allday = true
  end

  if mev.description
    gev.desc = NKF .nkf("-w", mev.description)
  else
    gev.desc = ""
  end

  gev.extended_property["MHC-Category"] = mev.category_as_string
  gev.extended_property["MHC-Record-Id"] = mev.rec_id

end

def gevents_to_mevents(mev, gev)

  #todo about google category tag

  if gev .title =~ /^(.*)\[([^\]]+)\]$/
    title = $1
  else
    title = gev.title
  end

  mev.set_subject(make_mail_header(title))
  mev.set_location(make_mail_header(gev.where))

  if gevents_allday?(gev)
    mev.set_time(nil, nil)
  elsif gev.st.nil?
    mev.set_time(nil, MhcTime .new(gev.en.strftime("%H:%M")))
#   elsif gev.en.nil?
#     mev.set_time(MhcTime .new(gev.st.strftime("%H:%M")), nil)
  else
    mev.set_time(MhcTime .new(gev.st.strftime("%H:%M")), 
                 MhcTime .new(gev.en.strftime("%H:%M")))
  end

  mev.set_category(gev.extended_property["MHC-Category"]) unless gev.extended_property["MHC-Category"].nil?
  mev.add_day(MhcDate .new(gev.st.strftime("%Y%m%d")))
  mev.set_description(NKF .nkf("-j", gev.desc)) unless gev.desc.nil?

end

def gevents_allday?(gev)

  return true if gev.allday

  # todo probably very rarely, start time to end time really a startTime to endTime is one day later and time 0:00
  #   or transparent?
  if gev.st.strftime("%H%M%S") == "000000"
    if gev.st + (24 * 60 * 60) == gev.en
      return true  
    end
  end

  return false

end

def make_mail_header(str)

  return nil if str.nil?

  ## todo clean code
  # append newline and any character let nkf append suffix
  tmp = NKF .nkf("-M", str + "\nA")

  return tmp[0, tmp.rindex("\n")]

end

@config = MhcGcalConfig.new(ARGV)

### upload EVENTS from MHC to Google Calendar ###

dbGCal = GCalDb .new(@config)
dbMhc = MhcDb .new(@config)

# delete google calendar entry when mhc entry was deleted.
dbMhc .deleted_entries .each do |dlog|

  gev = dbGCal .find_gevent(dlog .rec_id)
  next if gev .nil?
  print "Delete GCalendar entry: #{gev .to_s}\n" if $DEBUG
  gev.destroy!

end

# sync already associated entry or add google calendar entry
dbMhc .each_events do |date, mev, log|

#  p date, mev

  # todo find and compare update time or else
  gev = dbGCal .find_upload_event(mev, log)

  if gev
    print "Upload GCalendar: #{gev .to_s}\n" if $DEBUG
    mevents_to_gevents(gev, mev, date)
    gev.save!
  else
    #todo
#     gev = dbGCal .find_gevent(mev .rec_id)
#     if gev 
#       if gev.updated > log .mtime
#         dbMhc .add_sch(mev)
#       end
#     end
  end

end

# add mhc entry when google calendar entry was added
dbGCal .added_entries .each do |gev|

  mev = MhcScheduleItem.new()

  gevents_to_mevents(mev, gev)

  # !! sync key !!
  if gev.extended_property["MHC-Record-Id"].nil?
    gev.extended_property["MHC-Record-Id"] = mev.rec_id

    print "Save MHC-Record-Id to GCalendar: #{gev .to_s}\n" if $DEBUG
    # to save mhc Record-Id
    gev.save!
  end

  print "Add MHC entry: #{mev .to_s}\n" if $DEBUG
  dbMhc .add_sch(mev)

end


### Copyright Notice:

## Copyright (C) 2007 Masahiro Hayashi <mhayashi1120@gmail.com>. All rights reserved.

## Redistribution and use in source and binary forms, with or without
## modification, are permitted provided that the following conditions
## are met:
## 
## 1. Redistributions of source code must retain the above copyright
##    notice, this list of conditions and the following disclaimer.
## 2. Redistributions in binary form must reproduce the above copyright
##    notice, this list of conditions and the following disclaimer in the
##    documentation and/or other materials provided with the distribution.
## 3. Neither the name of the team nor the names of its contributors
##    may be used to endorse or promote products derived from this software
##    without specific prior written permission.
## 
## THIS SOFTWARE IS PROVIDED BY THE TEAM AND CONTRIBUTORS ``AS IS''
## AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
## LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
## FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL
## THE TEAM OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
## INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
## (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
## SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
## HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
## STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
## ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
## OF THE POSSIBILITY OF SUCH DAMAGE.

### mhc-gcal ends here
