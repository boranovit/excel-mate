class Column {
    constructor({ name, label, width, type, num, children, style, styleHeader, required }) {
        this.name = name || ""
        this.label = label || ""
        this.width = width || 20
        this.type = type || ""
        this.num = num
        this.children = children
        this.style = style
        this.styleHeader = styleHeader
        this.required = required
    }
}

module.exports = Column