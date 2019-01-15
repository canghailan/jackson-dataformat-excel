package cc.whohow.excel;

import com.fasterxml.jackson.core.FormatSchema;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonToken;
import com.fasterxml.jackson.core.PrettyPrinter;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.*;
import com.fasterxml.jackson.databind.introspect.BeanPropertyDefinition;

import java.io.IOException;
import java.lang.reflect.Type;
import java.util.TimeZone;
import java.util.function.Function;

public class ExcelMapper extends ObjectMapper {
    public ExcelMapper() {
        this(new ExcelFactory());
    }

    public ExcelMapper(ExcelFactory factory) {
        super(factory);
        setTimeZone(TimeZone.getDefault());
        setDateFormat(new ISO8601VariantDateFormat());
    }

    public ExcelMapper(ExcelMapper mapper) {
        super(mapper);
    }

    @Override
    public ExcelMapper copy() {
        _checkInvalidCopy(ExcelMapper.class);
        return new ExcelMapper(this);
    }

    @Override
    public ExcelFactory getFactory() {
        return (ExcelFactory) _jsonFactory;
    }

    @Override
    protected JsonToken _initForReading(JsonParser p, JavaType targetType) throws IOException {
        if (p.getSchema() == null) {
            p.setSchema(schemaForReader(targetType));
        }
        return super._initForReading(p, targetType);
    }

    @Override
    protected ObjectReader _newReader(DeserializationConfig config, JavaType valueType, Object valueToUpdate, FormatSchema schema, InjectableValues injectableValues) {
        if (schema == null) {
            schema = schemaFor(valueType, config::introspect);
        }
        return super._newReader(config, valueType, valueToUpdate, schema, injectableValues);
    }

    @Override
    protected ObjectWriter _newWriter(SerializationConfig config, JavaType rootType, PrettyPrinter pp) {
        return super._newWriter(config, rootType, pp)
                .with(schemaFor(rootType, config::introspect));
    }

    public ExcelSchema schemaForReader(Type type) {
        return schemaFor(getTypeFactory().constructType(type), getDeserializationConfig()::introspect);
    }

    public ExcelSchema schemaForReader(TypeReference<?> type) {
        return schemaFor(getTypeFactory().constructType(type), getDeserializationConfig()::introspect);
    }

    public ExcelSchema schemaForReader(JavaType type) {
        return schemaFor(type, getDeserializationConfig()::introspect);
    }

    public ExcelSchema schemaForWriter(Type type) {
        return schemaFor(getTypeFactory().constructType(type), getSerializationConfig()::introspect);
    }

    public ExcelSchema schemaForWriter(TypeReference<?> type) {
        return schemaFor(getTypeFactory().constructType(type), getSerializationConfig()::introspect);
    }

    public ExcelSchema schemaForWriter(JavaType type) {
        return schemaFor(type, getSerializationConfig()::introspect);
    }

    public ExcelSchema schemaFor(JavaType type, Function<JavaType, ? extends BeanDescription> introspect) {
        if (type == null) {
            return new ExcelSchema();
        }
        JavaType dataType = type;
        if (type.isArrayType() || type.isCollectionLikeType()) {
            dataType = type.getContentType();
        }
        return schemaFor(introspect.apply(dataType));
    }

    public ExcelSchema schemaFor(BeanDescription beanDescription) {
        ExcelSchema schema = new ExcelSchema();
        schema.withSheet(beanDescription.findClassDescription());

        for (BeanPropertyDefinition prop : beanDescription.findProperties()) {
            String name = prop.getName();
            String description = prop.getMetadata().getDescription();
            Integer index = prop.getMetadata().getIndex();
            schema.addKey(name, description == null ? name : description, index == null ? -1 : index);
        }
        return schema;
    }
}
