package cc.whohow.excel;

import java.text.DateFormat;
import java.text.FieldPosition;
import java.text.ParseException;
import java.text.ParsePosition;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.TimeZone;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ISO8601VariantDateFormat extends DateFormat {
    protected static final Pattern DATE_TIME = Pattern.compile(
            "(?<year>\\d{4}\\D+)?" +
                    "(?<month>\\d{1,2}|\\d{2}\\D+)" +
                    "(?<day>\\d{1,2}|\\d{2})" +
                    "(?<hour>\\D+\\d{1,2}|\\d{2}\\D+)?" +
                    "(?<minute>\\d{1,2}|\\d{2}\\D+)?" +
                    "(?<second>\\d{1,2}|\\d{2}\\D+)?" +
                    "(?<millisecond>\\d{1,3}|\\d{3})?");
    protected static final Pattern INT = Pattern.compile("\\d+");

    protected final ZoneId zoneId;

    public ISO8601VariantDateFormat() {
        this(ZoneId.systemDefault());
    }

    public ISO8601VariantDateFormat(TimeZone timeZone) {
        this(timeZone.toZoneId());
    }

    public ISO8601VariantDateFormat(ZoneId zoneId) {
        this.zoneId = zoneId;
    }

    @Override
    public void setTimeZone(TimeZone zone) {
        throw new UnsupportedOperationException();
    }

    @Override
    public TimeZone getTimeZone() {
        return TimeZone.getTimeZone(zoneId);
    }

    @Override
    public StringBuffer format(Date date, StringBuffer buffer, FieldPosition fieldPosition) {
        fieldPosition.setBeginIndex(0);
        fieldPosition.setEndIndex(0);

        LocalDateTime dateTime = LocalDateTime.ofInstant(date.toInstant(), zoneId);
        int year = dateTime.getYear();
        int month = dateTime.getMonthValue();
        int day = dateTime.getDayOfMonth();
        int hour = dateTime.getHour();
        int minute = dateTime.getMinute();
        int second = dateTime.getSecond();

        buffer.append(year);
        fieldPosition.setBeginIndex(0);
        fieldPosition.setEndIndex(4);

        buffer.append('-');
        format00(buffer, month);
        fieldPosition.setBeginIndex(5);
        fieldPosition.setEndIndex(7);

        buffer.append('-');
        format00(buffer, day);
        fieldPosition.setBeginIndex(8);
        fieldPosition.setEndIndex(10);

        if (hour == 0 && minute == 0 && second == 0) {
            return buffer;
        }

        buffer.append(' ');
        format00(buffer, hour);
        fieldPosition.setBeginIndex(11);
        fieldPosition.setEndIndex(13);

        buffer.append(':');
        format00(buffer, minute);
        fieldPosition.setBeginIndex(14);
        fieldPosition.setEndIndex(16);

        buffer.append(':');
        format00(buffer, second);
        fieldPosition.setBeginIndex(17);
        fieldPosition.setEndIndex(19);
        return buffer;
    }

    protected StringBuffer format00(StringBuffer buffer, int value) {
        return value < 10 ? buffer.append(0).append(value) : buffer.append(value);
    }

    @Override
    public Date parse(String source) throws ParseException {
        return super.parse(source);
    }

    @Override
    public Date parse(String source, ParsePosition pos) {
        if (pos.getIndex() != 0) {
            source = source.substring(pos.getIndex());
        }
        Matcher matcher = DATE_TIME.matcher(source);
        if (matcher.find()) {
            String year = matcher.group("year");
            String month = matcher.group("month");
            String day = matcher.group("day");
            String hour = matcher.group("hour");
            String minute = matcher.group("minute");
            String second = matcher.group("second");
            String millisecond = matcher.group("millisecond");
            return Date.from(LocalDateTime.of( //
                    year == null ? LocalDate.now().getYear() : Integer.parseInt(year), // 年，默认当前年
                    parseInt(month, 1), // 月
                    parseInt(day, 1), // 日
                    parseInt(hour, 0), // 时
                    parseInt(minute, 0), // 分
                    parseInt(second, 0), // 秒
                    parseInt(millisecond, 0) * 1_000_000) // 纳秒
                    .atZone(zoneId).toInstant());
        } else {
            pos.setErrorIndex(pos.getIndex());
            return null;
        }
    }

    private int parseInt(CharSequence text, int defaultValue) {
        if (text == null) {
            return defaultValue;
        }
        Matcher matcher = INT.matcher(text);
        if (matcher.find()) {
            return Integer.parseInt(matcher.group());
        }
        return defaultValue;
    }
}
