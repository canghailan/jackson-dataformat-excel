package cc.whohow.excel;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.util.Date;

public class ISO8601VariantDateSerializer extends StdSerializer<Date> {
    protected final ISO8601VariantDateFormat dateFormat;

    protected ISO8601VariantDateSerializer() {
        this(new ISO8601VariantDateFormat());
    }

    protected ISO8601VariantDateSerializer(ISO8601VariantDateFormat dateFormat) {
        super(Date.class);
        this.dateFormat = dateFormat;
    }

    @Override
    public void serialize(Date value, JsonGenerator gen, SerializerProvider provider) throws IOException {
        gen.writeString((value == null) ? null : dateFormat.format(value));
    }
}
