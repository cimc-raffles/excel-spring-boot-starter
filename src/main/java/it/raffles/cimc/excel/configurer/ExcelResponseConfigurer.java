package it.raffles.cimc.excel.configurer;

import java.util.List;
import java.util.stream.Collectors;

import org.springframework.beans.factory.InitializingBean;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.servlet.mvc.method.annotation.RequestMappingHandlerAdapter;
import org.springframework.web.servlet.mvc.method.annotation.RequestResponseBodyMethodProcessor;

import it.raffles.cimc.excel.interceptor.ExcelResponseHandler;


@Configuration
public class ExcelResponseConfigurer implements InitializingBean {

	@Autowired
	RequestMappingHandlerAdapter requestMappingHandlerAdapter;

	@Override
	public void afterPropertiesSet() throws Exception {
		List<HandlerMethodReturnValueHandler> handlers = requestMappingHandlerAdapter.getReturnValueHandlers();
		List<HandlerMethodReturnValueHandler> newHandlers = handlers.stream()
				.map(handler -> handler instanceof RequestResponseBodyMethodProcessor
						? new ExcelResponseHandler((RequestResponseBodyMethodProcessor)handler)
						: handler)
				.collect(Collectors.toList());

		requestMappingHandlerAdapter.setReturnValueHandlers(newHandlers);
	}

}