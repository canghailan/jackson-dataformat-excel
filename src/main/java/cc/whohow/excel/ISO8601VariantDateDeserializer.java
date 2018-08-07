package cc.whohow.excel;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.deser.std.StdDeserializer;

import java.io.IOException;
import java.text.ParseException;
import java.util.Date;

public class ISO8601VariantDateDeserializer extends StdDeserializer<Date> {
    protected final ISO8601VariantDateFormat dateFormat;

    protected ISO8601VariantDateDeserializer(ISO8601VariantDateFormat dateFormat) {
        super(Date.class);
        this.dateFormat = dateFormat;
    }

    @Override
    public Date deserialize(JsonParser p, DeserializationContext context) throws IOException, JsonProcessingException {
        try {
            String value = p.getValueAsString();
            return (value == null || value.isEmpty()) ? null : dateFormat.parse(value);
        } catch (ParseException e) {
            throw new IOException(e);
        }
    }
}
