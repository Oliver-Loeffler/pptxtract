module net.raumzeitfalle.pptxtract {
    requires java.logging;
    requires java.xml;
    
    requires info.picocli;
    
    opens net.raumzeitfalle.pptxtract to info.picocli;
}
