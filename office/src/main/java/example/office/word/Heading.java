package example.office.word;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

import java.util.ArrayList;
import java.util.List;

@Setter
@Getter
@ToString
@NoArgsConstructor
public class Heading {
    private String heading;
    private String text;
    // Title
    // TOC1、TOC2、TOC3
    // Heading1、Heading2、Heading3
    // Caption
    // fs-4
    // fs-4-first-line-indent-2
    private List<Heading> children;

    public Heading(String heading, String text) {
        this.heading = heading;
        this.text = text;
    }

    public void addChild(Heading heading) {
        if(children == null) {
            children = new ArrayList<>();
        }
        children.add(heading);
    }
    public Heading lastChild() {
        if(children == null || children.isEmpty()) {
            return null;
        }
        return children.get(children.size() - 1);
    }
}
