# public-YamlConverterVBA
Converts small segments of content between Yaml and MS Scripting Dictionary. 

## Purpose
Handling larger datasets with HP UFT / Microfocus UFT test automation which uses VBA.

Inspired by: https://stackoverflow.com/questions/38738162/yaml-parser-for-excel-vba

## Supported Examples

### Mapping Scalars to quoted Scalars, either separated either by new line or commas	
```
	{ key1: 'value1' 
	  key2: 'value2' }

	{ key1: 'value1' , key2: 'value2' }
```

### Mapping of Mappings
```
	testParent: { key1: 'value1', key2: 'value2' }
```

### Mapping Scalars to Scalars, either separated either by new line or commas	
```
	key1: value1 , key2: value2
	
	key1: value1
	key2: value2 
```

### Ignores Comments	
```
	key1: value1 # comment
	key2: value2 
```
