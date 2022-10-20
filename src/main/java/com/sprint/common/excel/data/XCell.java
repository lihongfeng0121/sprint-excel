/*
 * Copyright (C) GSX Techedu Inc. All Rights Reserved Unauthorized copying of this file, via any medium is strictly
 * prohibited Proprietary and confidential
 */
package com.sprint.common.excel.data;

import java.lang.annotation.*;

/**
 * 可导出的工具
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2020年02月18日
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface XCell {

    String name() default "";

    int sort() default -1;

    boolean titleUnfold() default false;

    boolean dataUnfold() default false;
}
