package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;

@Entity
@Table(name = "fx_setting")
@Getter @Setter
public class Setting {
    @Id
    @Column(name = "skey", length = 50)
    private String key;

    @Column(name = "sval", length = 500)
    private String value;

    public Setting() {}
    public Setting(String key, String value) { this.key = key; this.value = value; }
}
