package cc.whohow.excel;

import java.text.DateFormat;
import java.text.FieldPosition;
import java.text.ParsePosition;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;
import java.util.TimeZone;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ISO8601VariantDateFormat extends DateFormat {
    protected static final Pattern DATE_TIME = Pattern.compile(
            "(?<year>(?<Y>\\d{4}))?" +
                    "(?<month>\\D+(?<M1>\\d{1,2})|(?<M2>\\d{2}))" +
                    "(?<day>\\D+(?<d1>\\d{1,2})|(?<d2>\\d{2}))" +
                    "(?<hour>\\D+(?<H1>\\d{1,2})|(?<H2>\\d{2}))?" +
                    "(?<minute>\\D+(?<m1>\\d{1,2})|(?<m2>\\d{2}))?" +
                    "(?<second>\\D+(?<s1>\\d{1,2})|(?<s2>\\d{2}))?" +
                    "(?<millisecond>\\D+(?<S1>\\d{1,3})|(?<S2>\\d{3}))?");

    protected ZoneId zoneId;

    public ISO8601VariantDateFormat() {
        this(ZoneId.systemDefault());
    }

    public ISO8601VariantDateFormat(TimeZone timeZone) {
        this(timeZone.toZoneId());
    }

    public ISO8601VariantDateFormat(ZoneId zoneId) {
        this.zoneId = zoneId;
    }

    private static int parseYear(Matcher matcher, String group) {
        String value = matcher.group(group);
        if (value == null) {
            return LocalDate.now().getYear();
        }
        return Integer.parseInt(value);
    }

    private static int parsePart(Matcher matcher, String group1, String group2, int defaultValue) {
        String value1 = matcher.group(group1);
        if (value1 != null) {
            return Integer.parseInt(value1);
        }
        String value2 = matcher.group(group2);
        if (value2 != null) {
            return Integer.parseInt(value2);
        }
        return defaultValue;
    }

    @Override
    public TimeZone getTimeZone() {
        return TimeZone.getTimeZone(zoneId);
    }

    @Override
    public void setTimeZone(TimeZone zone) {
        this.zoneId = zone.toZoneId();
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
    public Date parse(String source, ParsePosition pos) {
        if (pos.getIndex() != 0) {
            source = source.substring(pos.getIndex());
        }
        Matcher matcher = DATE_TIME.matcher(source);
        if (matcher.find()) {
            Date date = Date.from(ZonedDateTime.of(//
                    parseYear(matcher, "Y"), // 年，默认当前年
                    parsePart(matcher, "M1", "M2", 1), // 月
                    parsePart(matcher, "d1", "d2", 1), // 日
                    parsePart(matcher, "H1", "H2", 0), // 时
                    parsePart(matcher, "m1", "m2", 0), // 分
                    parsePart(matcher, "s1", "s2", 0), // 秒
                    (int) TimeUnit.MILLISECONDS.toNanos(parsePart(matcher, "S1", "S2", 0)), // 毫秒
                    zoneId).toInstant());
            pos.setIndex(pos.getIndex() + matcher.group().length());
            return date;
        } else {
            pos.setErrorIndex(pos.getIndex());
            return null;
        }
    }

    @Override
    public Object clone() {
        return new ISO8601VariantDateFormat(zoneId);
    }
}
